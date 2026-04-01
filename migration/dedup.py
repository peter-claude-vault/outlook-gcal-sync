#!/usr/bin/env python3
"""
Deduplicate events that exist in both Exchange and Google calendars
(both email addresses invited to the same meeting).

Rules:
1. Teams meeting link → Exchange wins (delete Google copy)
2. Google Meet link → Google wins (delete Exchange synced copy)
3. Fallback: L'Oreal organizer domain → Exchange wins; Artefact domain → Google wins

Actions:
- Delete losing copy from Apple Calendar (EventKit)
- Output Google Calendar event IDs to delete via MCP
- Update sync_state.json exclusion list so sync doesn't re-create deleted events
"""

import json
import sys
import time
import hashlib
import re
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

SCRIPT_DIR = Path(__file__).parent
STATE_FILE = SCRIPT_DIR / "sync_state.json"
EXCLUDE_FILE = SCRIPT_DIR / "sync_exclusions.json"

LOREAL_DOMAINS = ["loreal.com", "lorealusa.com", "loreal.net"]
ARTEFACT_DOMAINS = ["artefact.com"]

# Time tolerance for matching (seconds) — events within 5 min are "same time"
TIME_TOLERANCE = 300


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
        print("Calendar access denied.")
        sys.exit(1)
    return store


def find_calendar(store, title):
    calendars = store.calendarsForEntityType_(0)
    for cal in calendars:
        if str(cal.title()) == title:
            return cal
    print(f"Calendar '{title}' not found.")
    sys.exit(1)


def get_events_detailed(store, calendar, days_forward=30):
    """Get events with full detail including attendees and description."""
    start = NSDate.date()
    end = NSDate.dateWithTimeIntervalSinceNow_(days_forward * 86400)

    predicate = store.predicateForEventsWithStartDate_endDate_calendars_(
        start, end, [calendar]
    )
    ek_events = store.eventsMatchingPredicate_(predicate)
    if not ek_events:
        return []

    events = []
    for ev in ek_events:
        title = str(ev.title()) if ev.title() else "(No title)"
        notes = str(ev.notes()) if ev.notes() else ""
        location = str(ev.location()) if ev.location() else ""
        start_ts = ev.startDate().timeIntervalSince1970()
        end_ts = ev.endDate().timeIntervalSince1970()
        all_day = bool(ev.isAllDay())
        ek_id = str(ev.eventIdentifier())

        # Get organizer email
        organizer_email = ""
        if ev.organizer():
            url = ev.organizer().URL()
            if url:
                spec = url.resourceSpecifier()
                if spec:
                    organizer_email = str(spec).replace("//", "")

        # Get attendee emails
        attendee_emails = []
        if ev.attendees():
            for att in ev.attendees():
                url = att.URL()
                if url:
                    spec = url.resourceSpecifier()
                    if spec:
                        attendee_emails.append(str(spec).replace("//", ""))

        # Detect meeting type from notes + location
        all_text = f"{notes} {location}".lower()
        has_teams = "teams.microsoft.com" in all_text or "microsoft teams" in all_text
        has_meet = "meet.google.com" in all_text

        events.append({
            "ek_id": ek_id,
            "title": title,
            "notes": notes,
            "location": location,
            "start_ts": start_ts,
            "end_ts": end_ts,
            "all_day": all_day,
            "organizer_email": organizer_email,
            "attendee_emails": attendee_emails,
            "has_teams": has_teams,
            "has_meet": has_meet,
            "event_ref": ev,
        })

    return events


def title_similar(a, b):
    """Check if two event titles are similar enough to be the same meeting."""
    # Normalize
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
    """
    Decide which calendar wins for a duplicate pair.
    Returns 'exchange' or 'google'.
    """
    # Rule 1: Teams → Exchange wins
    if exchange_ev["has_teams"] or google_ev["has_teams"]:
        return "exchange"

    # Rule 2: Google Meet → Google wins
    if exchange_ev["has_meet"] or google_ev["has_meet"]:
        return "google"

    # Rule 3: Organizer domain fallback
    # Check both events' organizer (should be same person)
    org_email = exchange_ev["organizer_email"] or google_ev["organizer_email"]
    domain = get_organizer_domain(org_email)

    if any(d in domain for d in LOREAL_DOMAINS):
        return "exchange"
    if any(d in domain for d in ARTEFACT_DOMAINS):
        return "google"

    # Default to exchange if unclear
    return "exchange"


def find_duplicates(exchange_events, google_events):
    """Find events that appear in both calendars."""
    duplicates = []
    google_matched = set()

    for ex_ev in exchange_events:
        for i, g_ev in enumerate(google_events):
            if i in google_matched:
                continue

            # Skip our own synced events
            if "[outlook-sync]" in g_ev.get("notes", ""):
                continue

            # Check time match
            time_diff = abs(ex_ev["start_ts"] - g_ev["start_ts"])
            if time_diff > TIME_TOLERANCE:
                continue

            # Check title match
            if not title_similar(ex_ev["title"], g_ev["title"]):
                continue

            google_matched.add(i)
            duplicates.append((ex_ev, g_ev))
            break

    return duplicates


def run_dedup(dry_run=True):
    config = json.loads((Path(__file__).parent.parent / "config.json").read_text())
    store = get_event_store()
    exchange_cal = find_calendar(store, config["source_calendar"])
    google_cal = find_calendar(store, config["target_calendar"])

    print("Reading Exchange calendar...")
    exchange_events = get_events_detailed(store, exchange_cal)
    print(f"  {len(exchange_events)} events")

    print("Reading Google calendar...")
    google_events = get_events_detailed(store, google_cal)
    # Filter out [outlook-sync] events for matching purposes
    native_google = [e for e in google_events if "[outlook-sync]" not in e.get("notes", "")]
    synced_google = [e for e in google_events if "[outlook-sync]" in e.get("notes", "")]
    print(f"  {len(native_google)} native events, {len(synced_google)} synced events")

    print("\nFinding duplicates...")
    duplicates = find_duplicates(exchange_events, native_google)
    print(f"  Found {len(duplicates)} duplicate pairs\n")

    # Load sync state to find synced copies
    state = json.loads(STATE_FILE.read_text()) if STATE_FILE.exists() else {}
    exclusions = json.loads(EXCLUDE_FILE.read_text()) if EXCLUDE_FILE.exists() else []

    exchange_deletes = []  # EK events to delete from Exchange calendar in Apple Calendar
    google_native_deletes = []  # EK events to delete from Google calendar in Apple Calendar
    google_synced_deletes = []  # Synced copies to delete from Google calendar
    google_mcp_deletes = []  # Event summaries+times for MCP deletion
    new_exclusions = []  # Exchange event IDs to exclude from future syncs

    for ex_ev, g_ev in duplicates:
        winner = decide_winner(ex_ev, g_ev)
        loser_label = "Google" if winner == "exchange" else "Exchange"

        meeting_type = ""
        if ex_ev["has_teams"] or g_ev["has_teams"]:
            meeting_type = "Teams"
        elif ex_ev["has_meet"] or g_ev["has_meet"]:
            meeting_type = "Meet"
        else:
            meeting_type = f"org: {get_organizer_domain(ex_ev['organizer_email'] or g_ev['organizer_email'])}"

        print(f"  {ex_ev['title']}")
        print(f"    → Winner: {winner.upper()} ({meeting_type}), deleting {loser_label} copy")

        if winner == "exchange":
            # Delete native Google copy (in Apple Calendar)
            google_native_deletes.append(g_ev)
            # Keep the [outlook-sync] copy on Google (or it'll be re-synced)
        else:
            # Google wins — delete Exchange copy from Apple Calendar is optional
            # (it's the native Exchange invite, not harmful in Apple Calendar)
            # But we DO need to:
            # 1. Delete the [outlook-sync] copy from Google calendar
            # 2. Exclude this Exchange event from future syncs
            ex_id = ex_ev["ek_id"]
            new_exclusions.append(ex_id)

            # Find and mark the synced copy for deletion
            for s_ev in synced_google:
                if ex_id in s_ev.get("notes", ""):
                    google_synced_deletes.append(s_ev)
                    break

            # Also find matching synced copy in state
            if ex_id in state:
                google_synced_deletes_ids = state[ex_id].get("target_ek_id")

    if dry_run:
        print(f"\n--- DRY RUN SUMMARY ---")
        print(f"  Native Google events to delete from Apple Calendar: {len(google_native_deletes)}")
        print(f"  Synced copies to delete from Google calendar: {len(google_synced_deletes)}")
        print(f"  Exchange events to exclude from future syncs: {len(new_exclusions)}")
        print(f"\nRun with --apply to execute.")
        return

    # Execute deletions
    print(f"\n--- APPLYING CHANGES ---")

    # Delete native Google events from Apple Calendar
    deleted_native = 0
    for ev_data in google_native_deletes:
        ev = ev_data["event_ref"]
        success = store.removeEvent_span_error_(ev, EKSpanThisEvent, None)
        if success:
            deleted_native += 1
            print(f"  Deleted (Google native): {ev_data['title']}")
        else:
            print(f"  FAILED (Google native): {ev_data['title']}")

    # Delete synced copies from Google calendar
    deleted_synced = 0
    for ev_data in google_synced_deletes:
        ev = ev_data["event_ref"]
        success = store.removeEvent_span_error_(ev, EKSpanThisEvent, None)
        if success:
            deleted_synced += 1
            print(f"  Deleted (synced copy): {ev_data['title']}")
        else:
            print(f"  FAILED (synced copy): {ev_data['title']}")

    # Update sync state — remove excluded events
    for ex_id in new_exclusions:
        if ex_id in state:
            del state[ex_id]

    STATE_FILE.write_text(json.dumps(state, indent=2))

    # Save exclusions so sync script skips these
    existing_exclusions = json.loads(EXCLUDE_FILE.read_text()) if EXCLUDE_FILE.exists() else []
    all_exclusions = list(set(existing_exclusions + new_exclusions))
    EXCLUDE_FILE.write_text(json.dumps(all_exclusions, indent=2))

    print(f"\n--- DONE ---")
    print(f"  Deleted native Google events: {deleted_native}")
    print(f"  Deleted synced copies: {deleted_synced}")
    print(f"  Added sync exclusions: {len(new_exclusions)}")

    # Output info for Google Calendar MCP cleanup
    # These are the native Google events that need deleting via MCP too
    # (EventKit deletion propagates to Google, but flagging for verification)
    print(f"\n  EventKit deletions will propagate to Google Calendar automatically.")


if __name__ == "__main__":
    dry_run = "--apply" not in sys.argv
    run_dedup(dry_run=dry_run)
