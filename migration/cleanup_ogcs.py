#!/usr/bin/env python3
"""
Clean up OGCS-created duplicate events from Google Calendar via EventKit.

Run this ONCE before starting the sync if you're migrating from OGCS
(Outlook Google Calendar Sync). It removes events that have .ogcs markers
in their organizer, notes, or attendee fields.

Usage:
    # Dry run — shows what would be deleted
    python3 migration/cleanup_ogcs.py

    # Actually delete
    python3 migration/cleanup_ogcs.py --delete
"""

import json
import sys
import time
from pathlib import Path

import objc
from EventKit import (
    EKEventStore,
    EKEntityMaskEvent,
    EKSpanThisEvent,
    EKAuthorizationStatusFullAccess,
    EKAuthorizationStatusWriteOnly,
)
from Foundation import NSDate

SCRIPT_DIR = Path(__file__).parent.parent
CONFIG_FILE = SCRIPT_DIR / "config.json"


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
        print("Calendar access denied. Grant in System Settings → Privacy & Security → Calendars.")
        sys.exit(1)
    return store


def find_calendar(store, title):
    calendars = store.calendarsForEntityType_(0)
    for cal in calendars:
        if str(cal.title()) == title:
            return cal
    print(f"Calendar '{title}' not found. Available:")
    for cal in calendars:
        print(f"  - {cal.title()} (source: {cal.source().title() if cal.source() else '?'})")
    sys.exit(1)


def find_ogcs_events(store, calendar):
    """Find events with .ogcs markers."""
    start = NSDate.date()
    end = NSDate.dateWithTimeIntervalSinceNow_(30 * 86400)

    predicate = store.predicateForEventsWithStartDate_endDate_calendars_(
        start, end, [calendar]
    )
    events = store.eventsMatchingPredicate_(predicate)
    if not events:
        return [], []

    ogcs_events = []
    clean_events = []

    for ev in events:
        title = str(ev.title()) if ev.title() else "(No title)"
        notes = str(ev.notes()) if ev.notes() else ""
        organizer = ""
        if ev.organizer():
            url = ev.organizer().URL()
            if url and url.resourceSpecifier():
                organizer = str(url.resourceSpecifier())

        attendee_emails = []
        if ev.attendees():
            for att in ev.attendees():
                url = att.URL()
                if url and url.resourceSpecifier():
                    attendee_emails.append(str(url.resourceSpecifier()))

        all_text = f"{notes} {organizer} {' '.join(attendee_emails)}".lower()

        if ".ogcs" in all_text:
            ogcs_events.append({"title": title, "event_ref": ev})
        else:
            clean_events.append({"title": title})

    return ogcs_events, clean_events


def main():
    if not CONFIG_FILE.exists():
        print("Error: config.json not found. Run install.sh or sync.py --setup first.")
        sys.exit(1)

    config = json.loads(CONFIG_FILE.read_text())
    target_name = config["target_calendar"]

    store = get_event_store()
    cal = find_calendar(store, target_name)

    print(f"Scanning '{target_name}' for OGCS events (next 30 days)...")
    ogcs_events, clean_events = find_ogcs_events(store, cal)

    print(f"\n  OGCS events found: {len(ogcs_events)}")
    print(f"  Clean events: {len(clean_events)}")

    if not ogcs_events:
        print("\nNo OGCS events to clean up.")
        return

    for e in ogcs_events[:15]:
        print(f"    - {e['title']}")
    if len(ogcs_events) > 15:
        print(f"    ... and {len(ogcs_events) - 15} more")

    if "--delete" not in sys.argv:
        print(f"\nDry run. Run with --delete to remove {len(ogcs_events)} OGCS events.")
        return

    print(f"\nDeleting {len(ogcs_events)} OGCS events...")
    deleted, failed = 0, 0
    for e in ogcs_events:
        if store.removeEvent_span_error_(e["event_ref"], EKSpanThisEvent, None):
            deleted += 1
        else:
            failed += 1
            print(f"  FAILED: {e['title']}")

    print(f"\nDone. Deleted={deleted}, Failed={failed}")


if __name__ == "__main__":
    main()
