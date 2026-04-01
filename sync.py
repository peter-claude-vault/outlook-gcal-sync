#!/usr/bin/env python3
"""
Outlook → Google Calendar one-way sync via macOS Calendar (EventKit).

Reads events from an Exchange calendar and writes them to a Google calendar,
both accessed through macOS Calendar.app's EventKit. No external APIs needed —
both accounts are added via System Settings → Internet Accounts.

Includes automatic deduplication: when both email addresses are invited to the
same meeting, keeps one copy based on meeting type (Teams → Exchange, Meet →
Google) or organizer domain (L'Oreal → Exchange, Artefact → Google).
"""

import json
import re
import sys
import hashlib
import logging
import time
from datetime import datetime, timezone
from pathlib import Path
from difflib import SequenceMatcher

import objc
from EventKit import (
    EKEvent,
    EKEventStore,
    EKEntityMaskEvent,
    EKSpanThisEvent,
    EKAuthorizationStatusFullAccess,
    EKAuthorizationStatusWriteOnly,
)
from Foundation import NSDate

# --- Config ---
SCRIPT_DIR = Path(__file__).parent
STATE_FILE = SCRIPT_DIR / "sync_state.json"
EXCLUDE_FILE = SCRIPT_DIR / "sync_exclusions.json"
CONFIG_FILE = SCRIPT_DIR / "config.json"

SYNC_DAYS_BACK = 1
SYNC_DAYS_FORWARD = 30
SYNC_TAG = "[outlook-sync]"

# Dedup config
LOREAL_DOMAINS = ["loreal.com", "lorealusa.com", "loreal.net"]
ARTEFACT_DOMAINS = ["artefact.com"]
TIME_TOLERANCE = 300  # seconds — events within 5 min count as "same time"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("outlook-gcal-sync")


# ---------------------------------------------------------------------------
# EventKit helpers
# ---------------------------------------------------------------------------

def load_config():
    if not CONFIG_FILE.exists():
        log.error("config.json not found. Run: python3 sync.py --setup")
        sys.exit(1)
    return json.loads(CONFIG_FILE.read_text())


def get_event_store():
    store = EKEventStore.alloc().init()
    status = EKEventStore.authorizationStatusForEntityType_(EKEntityMaskEvent)
    if status in (EKAuthorizationStatusFullAccess, EKAuthorizationStatusWriteOnly):
        return store
    granted = [None]
    def callback(g, e):
        granted[0] = g
    store.requestFullAccessToEventsWithCompletion_(callback)
    for _ in range(100):
        if granted[0] is not None:
            break
        time.sleep(0.1)
    if not granted[0]:
        log.error("Calendar access denied. Grant in System Settings → Privacy & Security → Calendars.")
        sys.exit(1)
    return store


def list_calendars(store):
    calendars = store.calendarsForEntityType_(0)
    result = []
    for cal in calendars:
        result.append({
            "title": str(cal.title()),
            "id": str(cal.calendarIdentifier()),
            "source": str(cal.source().title()) if cal.source() else "Unknown",
            "type": int(cal.type()),
        })
    return result


def find_calendar(store, title):
    calendars = store.calendarsForEntityType_(0)
    for cal in calendars:
        if str(cal.title()) == title:
            return cal
    log.error(f"Calendar '{title}' not found. Available:")
    for cal in calendars:
        log.error(f"  - {cal.title()} (source: {cal.source().title() if cal.source() else '?'})")
    sys.exit(1)


def get_events(store, calendar):
    """Read events from a calendar within the sync window (basic fields)."""
    start = NSDate.dateWithTimeIntervalSinceNow_(-SYNC_DAYS_BACK * 86400)
    end = NSDate.dateWithTimeIntervalSinceNow_(SYNC_DAYS_FORWARD * 86400)
    predicate = store.predicateForEventsWithStartDate_endDate_calendars_(start, end, [calendar])
    ek_events = store.eventsMatchingPredicate_(predicate)
    if ek_events is None:
        return []

    events = []
    for ev in ek_events:
        summary = str(ev.title()) if ev.title() else "(No title)"
        location = str(ev.location()) if ev.location() else ""
        notes = str(ev.notes()) if ev.notes() else ""
        start_ts = ev.startDate().timeIntervalSince1970()
        end_ts = ev.endDate().timeIntervalSince1970()
        all_day = bool(ev.isAllDay())

        content_hash = hashlib.sha256(
            f"{summary}|{location}|{start_ts}|{end_ts}|{all_day}".encode()
        ).hexdigest()[:16]

        # Use ek_id + start_ts to uniquely identify recurring event occurrences
        occurrence_id = f"{ev.eventIdentifier()}@{int(start_ts)}"

        events.append({
            "ek_id": occurrence_id,
            "base_ek_id": str(ev.eventIdentifier()),
            "summary": summary,
            "location": location,
            "notes": notes,
            "start_ts": start_ts,
            "end_ts": end_ts,
            "all_day": all_day,
            "content_hash": content_hash,
        })
    return events


def get_events_detailed(store, calendar):
    """Read events with organizer/attendee/meeting-type info for dedup."""
    start = NSDate.dateWithTimeIntervalSinceNow_(-SYNC_DAYS_BACK * 86400)
    end = NSDate.dateWithTimeIntervalSinceNow_(SYNC_DAYS_FORWARD * 86400)
    predicate = store.predicateForEventsWithStartDate_endDate_calendars_(start, end, [calendar])
    ek_events = store.eventsMatchingPredicate_(predicate)
    if ek_events is None:
        return []

    events = []
    for ev in ek_events:
        title = str(ev.title()) if ev.title() else "(No title)"
        notes = str(ev.notes()) if ev.notes() else ""
        location = str(ev.location()) if ev.location() else ""
        start_ts = ev.startDate().timeIntervalSince1970()
        base_ek_id = str(ev.eventIdentifier())
        occurrence_id = f"{base_ek_id}@{int(start_ts)}"

        organizer_email = ""
        if ev.organizer():
            url = ev.organizer().URL()
            if url and url.resourceSpecifier():
                organizer_email = str(url.resourceSpecifier()).replace("//", "")

        all_text = f"{notes} {location}".lower()

        events.append({
            "ek_id": occurrence_id,
            "base_ek_id": base_ek_id,
            "title": title,
            "notes": notes,
            "start_ts": start_ts,
            "organizer_email": organizer_email,
            "has_teams": "teams.microsoft.com" in all_text or "microsoft teams" in all_text,
            "has_meet": "meet.google.com" in all_text,
            "event_ref": ev,
        })
    return events


def create_ek_event(store, target_cal, ev):
    new_event = EKEvent.eventWithEventStore_(store)
    new_event.setTitle_(ev["summary"])
    new_event.setLocation_(ev["location"])
    new_event.setCalendar_(target_cal)
    new_event.setAllDay_(ev["all_day"])
    new_event.setStartDate_(NSDate.dateWithTimeIntervalSince1970_(ev["start_ts"]))
    new_event.setEndDate_(NSDate.dateWithTimeIntervalSince1970_(ev["end_ts"]))
    notes = ev["notes"]
    if notes:
        new_event.setNotes_(f"{SYNC_TAG} {ev['ek_id']}\n{notes}")
    else:
        new_event.setNotes_(f"{SYNC_TAG} {ev['ek_id']}")
    success = store.saveEvent_span_error_(new_event, EKSpanThisEvent, None)
    if success:
        return str(new_event.eventIdentifier())
    return None


def update_ek_event(store, target_cal, gcal_ek_id, ev):
    existing = store.eventWithIdentifier_(gcal_ek_id)
    if existing is None:
        return False
    existing.setTitle_(ev["summary"])
    existing.setLocation_(ev["location"])
    existing.setAllDay_(ev["all_day"])
    existing.setStartDate_(NSDate.dateWithTimeIntervalSince1970_(ev["start_ts"]))
    existing.setEndDate_(NSDate.dateWithTimeIntervalSince1970_(ev["end_ts"]))
    notes = ev["notes"]
    if notes:
        existing.setNotes_(f"{SYNC_TAG} {ev['ek_id']}\n{notes}")
    else:
        existing.setNotes_(f"{SYNC_TAG} {ev['ek_id']}")
    return bool(store.saveEvent_span_error_(existing, EKSpanThisEvent, None))


def delete_ek_event(store, gcal_ek_id):
    existing = store.eventWithIdentifier_(gcal_ek_id)
    if existing is None:
        return True
    return bool(store.removeEvent_span_error_(existing, EKSpanThisEvent, None))


# ---------------------------------------------------------------------------
# State / exclusion helpers
# ---------------------------------------------------------------------------

def load_state():
    if STATE_FILE.exists():
        return json.loads(STATE_FILE.read_text())
    return {}

def save_state(state):
    STATE_FILE.write_text(json.dumps(state, indent=2))

def load_exclusions():
    if EXCLUDE_FILE.exists():
        return set(json.loads(EXCLUDE_FILE.read_text()))
    return set()

def save_exclusions(exclusions):
    EXCLUDE_FILE.write_text(json.dumps(sorted(exclusions), indent=2))


# ---------------------------------------------------------------------------
# Dedup logic
# ---------------------------------------------------------------------------

def title_similar(a, b):
    a = re.sub(r"^(Canceled|Cancelled):\s*", "", a).strip().lower()
    b = re.sub(r"^(Canceled|Cancelled):\s*", "", b).strip().lower()
    if a == b:
        return True
    return SequenceMatcher(None, a, b).ratio() > 0.75


def get_organizer_domain(email):
    if "@" in email:
        return email.split("@")[1].lower()
    return ""


def decide_winner(exchange_ev, google_ev):
    """Artefact organizer → Google. Everyone else → Exchange."""
    org_email = exchange_ev["organizer_email"] or google_ev["organizer_email"]
    domain = get_organizer_domain(org_email)
    if any(d in domain for d in ARTEFACT_DOMAINS):
        return "google"
    return "exchange"


def dedup(store, source_cal, target_cal, state, exclusions):
    """
    Find events that exist natively in both calendars (both emails invited).
    Delete the losing copy and update exclusions so sync doesn't re-create.
    Returns updated (state, exclusions).
    """
    exchange_events = get_events_detailed(store, source_cal)
    google_events = get_events_detailed(store, target_cal)

    native_google = [e for e in google_events if SYNC_TAG not in e["notes"]]
    synced_google = [e for e in google_events if SYNC_TAG in e["notes"]]

    # Match duplicates
    google_matched = set()
    duplicates = []
    for ex_ev in exchange_events:
        for i, g_ev in enumerate(native_google):
            if i in google_matched:
                continue
            if abs(ex_ev["start_ts"] - g_ev["start_ts"]) > TIME_TOLERANCE:
                continue
            if not title_similar(ex_ev["title"], g_ev["title"]):
                continue
            google_matched.add(i)
            duplicates.append((ex_ev, g_ev))
            break

    if not duplicates:
        return state, exclusions

    google_native_deletes = 0
    synced_deletes = 0
    new_exclusions = 0

    for ex_ev, g_ev in duplicates:
        winner = decide_winner(ex_ev, g_ev)

        if winner == "exchange":
            # Delete native Google copy; synced copy stays
            ev_ref = g_ev["event_ref"]
            if store.removeEvent_span_error_(ev_ref, EKSpanThisEvent, None):
                google_native_deletes += 1
                log.info(f"  Dedup (Exchange wins): {ex_ev['title']}")
        else:
            # Google wins — exclude base ID from sync (covers all occurrences)
            base_id = ex_ev.get("base_ek_id", ex_ev["ek_id"])
            if base_id not in exclusions:
                exclusions.add(base_id)
                new_exclusions += 1

            # Remove any synced copies from Google calendar
            occ_id = ex_ev["ek_id"]
            if occ_id in state:
                target_ek_id = state[occ_id].get("target_ek_id")
                if target_ek_id:
                    delete_ek_event(store, target_ek_id)
                    synced_deletes += 1
                del state[occ_id]

            # Also check synced events by notes tag (using base ID)
            for s_ev in synced_google:
                if base_id in s_ev.get("notes", ""):
                    store.removeEvent_span_error_(s_ev["event_ref"], EKSpanThisEvent, None)
                    break

            log.info(f"  Dedup (Google wins): {g_ev['title']}")

    if google_native_deletes or synced_deletes or new_exclusions:
        log.info(f"  Dedup summary: native_deleted={google_native_deletes}, "
                 f"synced_deleted={synced_deletes}, new_exclusions={new_exclusions}")

    return state, exclusions


# ---------------------------------------------------------------------------
# Main sync
# ---------------------------------------------------------------------------

def sync(config):
    source_name = config["source_calendar"]
    target_name = config["target_calendar"]

    store = get_event_store()
    source_cal = find_calendar(store, source_name)
    target_cal = find_calendar(store, target_name)

    # --- Phase 1: Dedup ---
    log.info("Phase 1: Checking for cross-calendar duplicates...")
    state = load_state()
    exclusions = load_exclusions()
    state, exclusions = dedup(store, source_cal, target_cal, state, exclusions)
    save_exclusions(exclusions)

    # --- Phase 2: Sync Exchange → Google ---
    log.info(f"Phase 2: Syncing '{source_name}' → '{target_name}'...")
    events = get_events(store, source_cal)

    # Read target calendar for pre-create collision check.
    # EventKit dedup (Phase 1) can miss native Google events that aren't
    # visible via EventKit (e.g. recurring events after occurrence deletion).
    # This secondary check prevents creating synced duplicates.
    target_events = get_events(store, target_cal)
    native_target = [e for e in target_events if SYNC_TAG not in e.get("notes", "")]

    log.info(f"  {len(events)} source events, {len(exclusions)} exclusions")

    current_ids = set()
    created, updated, deleted, unchanged, skipped = 0, 0, 0, 0, 0

    for ev in events:
        ek_id = ev["ek_id"]
        current_ids.add(ek_id)

        base_id = ev.get("base_ek_id", ek_id)
        if ek_id in exclusions or base_id in exclusions:
            skipped += 1
            continue

        # Skip canceled events — and delete synced copy if one exists
        is_canceled = ev["summary"].lower().startswith("canceled:") or \
                      ev["summary"].lower().startswith("cancelled:")
        if is_canceled:
            if ek_id in state:
                delete_ek_event(store, state[ek_id]["target_ek_id"])
                del state[ek_id]
                deleted += 1
                log.info(f"  Removed canceled: {ev['summary']}")
            skipped += 1
            continue

        if ek_id in state:
            if state[ek_id]["content_hash"] == ev["content_hash"]:
                unchanged += 1
                continue
            if update_ek_event(store, target_cal, state[ek_id]["target_ek_id"], ev):
                state[ek_id]["content_hash"] = ev["content_hash"]
                updated += 1
                log.info(f"  Updated: {ev['summary']}")
            else:
                log.warning(f"  Failed to update: {ev['summary']}")
        else:
            # Pre-create collision check: skip if a native event with same
            # title and time already exists on the target calendar
            collision = False
            for t_ev in native_target:
                if abs(t_ev["start_ts"] - ev["start_ts"]) <= TIME_TOLERANCE \
                        and title_similar(t_ev["summary"], ev["summary"]):
                    collision = True
                    break
            if collision:
                exclusions.add(base_id)
                skipped += 1
                log.info(f"  Collision skip (native exists): {ev['summary']}")
                continue

            new_id = create_ek_event(store, target_cal, ev)
            if new_id:
                state[ek_id] = {
                    "target_ek_id": new_id,
                    "content_hash": ev["content_hash"],
                }
                created += 1
                log.info(f"  Created: {ev['summary']}")
            else:
                log.warning(f"  Failed to create: {ev['summary']}")

    # Delete events removed from Exchange (but not excluded ones)
    stale_ids = set(state.keys()) - current_ids - exclusions
    for ek_id in stale_ids:
        if delete_ek_event(store, state[ek_id]["target_ek_id"]):
            deleted += 1
            log.info(f"  Deleted synced event (removed from Outlook)")
        else:
            log.warning(f"  Failed to delete stale event")
        del state[ek_id]

    save_state(state)
    log.info(f"Done. Created={created} Updated={updated} Deleted={deleted} "
             f"Unchanged={unchanged} Skipped={skipped}")


# ---------------------------------------------------------------------------
# Setup wizard
# ---------------------------------------------------------------------------

def setup_wizard():
    print("\n=== Outlook → Google Calendar Sync Setup ===\n")
    store = get_event_store()
    calendars = list_calendars(store)

    print("Calendars found in macOS Calendar:\n")
    for i, cal in enumerate(calendars):
        marker = ""
        if "exchange" in cal["source"].lower() or cal["type"] == 2:
            marker = " ← Exchange"
        elif "google" in cal["source"].lower() or "gmail" in cal["source"].lower():
            marker = " ← Google"
        print(f"  [{i}] {cal['title']}  (source: {cal['source']}){marker}")

    print()
    choice = input("Enter the number of your L'Oreal (source) calendar: ").strip()
    try:
        source = calendars[int(choice)]
    except (ValueError, IndexError):
        print("Invalid choice.")
        sys.exit(1)

    print()
    choice = input("Enter the number of your Google (target) calendar: ").strip()
    try:
        target = calendars[int(choice)]
    except (ValueError, IndexError):
        print("Invalid choice.")
        sys.exit(1)

    print(f"\nSource: {source['title']} ({source['source']})")
    print(f"Target: {target['title']} ({target['source']})")

    config = {
        "source_calendar": source["title"],
        "target_calendar": target["title"],
    }
    CONFIG_FILE.write_text(json.dumps(config, indent=2))
    print(f"\nConfig saved to {CONFIG_FILE}")
    print(f"\nSetup complete. Run 'python3 {__file__}' to sync.")


if __name__ == "__main__":
    if "--setup" in sys.argv:
        setup_wizard()
    elif "--list-calendars" in sys.argv:
        store = get_event_store()
        for cal in list_calendars(store):
            print(f"  {cal['title']}  (source: {cal['source']}, type: {cal['type']})")
    else:
        config = load_config()
        sync(config)
