#!/usr/bin/env python3
"""
Clean up OGCS-created duplicate events from Google Calendar via EventKit.
Removes events from today forward that have .OGCS in organizer/notes/attendees.
"""

import sys
import time
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


def inspect_events(store, calendar):
    """Print details of events to understand OGCS markers."""
    start = NSDate.date()  # today
    end = NSDate.dateWithTimeIntervalSinceNow_(30 * 86400)

    predicate = store.predicateForEventsWithStartDate_endDate_calendars_(
        start, end, [calendar]
    )
    events = store.eventsMatchingPredicate_(predicate)
    if not events:
        print("No events found.")
        return

    ogcs_events = []
    clean_events = []

    for ev in events:
        title = str(ev.title()) if ev.title() else "(No title)"
        notes = str(ev.notes()) if ev.notes() else ""
        organizer = ""
        if ev.organizer():
            url = ev.organizer().URL()
            if url:
                organizer = str(url.resourceSpecifier()) if url.resourceSpecifier() else str(url)

        # Check attendees too
        attendee_emails = []
        if ev.attendees():
            for att in ev.attendees():
                url = att.URL()
                if url:
                    email = str(url.resourceSpecifier()) if url.resourceSpecifier() else str(url)
                    attendee_emails.append(email)

        all_text = f"{notes} {organizer} {' '.join(attendee_emails)}".lower()
        is_ogcs = ".ogcs" in all_text

        entry = {
            "title": title,
            "organizer": organizer,
            "notes_preview": notes[:100],
            "attendees_sample": attendee_emails[:3],
            "is_ogcs": is_ogcs,
            "event_ref": ev,
        }

        if is_ogcs:
            ogcs_events.append(entry)
        else:
            clean_events.append(entry)

    print(f"\n=== OGCS events (to delete): {len(ogcs_events)} ===")
    for e in ogcs_events[:10]:
        print(f"  {e['title']}")
        if e['organizer']:
            print(f"    organizer: {e['organizer']}")
        if e['notes_preview']:
            print(f"    notes: {e['notes_preview']}")

    if len(ogcs_events) > 10:
        print(f"  ... and {len(ogcs_events) - 10} more")

    print(f"\n=== Clean events (keeping): {len(clean_events)} ===")
    for e in clean_events[:10]:
        marker = ""
        if "[outlook-sync]" in (e.get("notes_preview") or ""):
            marker = " [our sync]"
        print(f"  {e['title']}{marker}")
        if e['organizer']:
            print(f"    organizer: {e['organizer']}")

    if len(clean_events) > 10:
        print(f"  ... and {len(clean_events) - 10} more")

    return ogcs_events, clean_events


def delete_ogcs_events(store, calendar):
    """Delete all OGCS events from today forward."""
    start = NSDate.date()
    end = NSDate.dateWithTimeIntervalSinceNow_(30 * 86400)

    predicate = store.predicateForEventsWithStartDate_endDate_calendars_(
        start, end, [calendar]
    )
    events = store.eventsMatchingPredicate_(predicate)
    if not events:
        print("No events found.")
        return

    deleted = 0
    failed = 0

    for ev in events:
        notes = str(ev.notes()) if ev.notes() else ""
        organizer = ""
        if ev.organizer():
            url = ev.organizer().URL()
            if url:
                organizer = str(url.resourceSpecifier()) if url.resourceSpecifier() else str(url)

        attendee_emails = []
        if ev.attendees():
            for att in ev.attendees():
                url = att.URL()
                if url:
                    email = str(url.resourceSpecifier()) if url.resourceSpecifier() else str(url)
                    attendee_emails.append(email)

        all_text = f"{notes} {organizer} {' '.join(attendee_emails)}".lower()

        if ".ogcs" in all_text:
            title = str(ev.title()) if ev.title() else "(No title)"
            success = store.removeEvent_span_error_(ev, EKSpanThisEvent, None)
            if success:
                deleted += 1
                print(f"  Deleted: {title}")
            else:
                failed += 1
                print(f"  FAILED: {title}")

    print(f"\nDone. Deleted={deleted}, Failed={failed}")


if __name__ == "__main__":
    import json
    from pathlib import Path
    config = json.loads((Path(__file__).parent.parent / "config.json").read_text())
    store = get_event_store()
    cal = find_calendar(store, config["target_calendar"])

    if "--delete" in sys.argv:
        delete_ogcs_events(store, cal)
    else:
        inspect_events(store, cal)
        print("\nRun with --delete to remove OGCS events.")
