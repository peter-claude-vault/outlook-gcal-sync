#!/usr/bin/env python3
"""
One-shot retroactive pass: 14 days back + 30 days forward.
Applies dedup rules and canceled-event removal across the full window.
Also cleans canceled events from native Google calendar.
"""
import sync
import re
from EventKit import EKSpanThisEvent

# Expand window to cover past 2 weeks + full forward range
sync.SYNC_DAYS_BACK = 14
sync.SYNC_DAYS_FORWARD = 30

store = sync.get_event_store()
config = sync.load_config()
source_cal = sync.find_calendar(store, config["source_calendar"])
target_cal = sync.find_calendar(store, config["target_calendar"])

# --- Phase 1: Dedup with new rules ---
sync.log.info("=== Phase 1: Dedup (14 days back + 30 forward) ===")
state = sync.load_state()
exclusions = sync.load_exclusions()
state, exclusions = sync.dedup(store, source_cal, target_cal, state, exclusions)
sync.save_exclusions(exclusions)

# --- Phase 2: Remove canceled synced events ---
sync.log.info("=== Phase 2: Remove canceled synced events ===")
events = sync.get_events(store, source_cal)

canceled_removed = 0
for ev in events:
    is_canceled = ev["summary"].lower().startswith("canceled:") or \
                  ev["summary"].lower().startswith("cancelled:")
    if is_canceled and ev["ek_id"] in state:
        sync.delete_ek_event(store, state[ev["ek_id"]]["target_ek_id"])
        del state[ev["ek_id"]]
        canceled_removed += 1
        sync.log.info(f"  Removed: {ev['summary']}")

sync.log.info(f"  Canceled synced events removed: {canceled_removed}")

# --- Phase 3: Remove canceled native Google events ---
sync.log.info("=== Phase 3: Remove canceled native Google events ===")
google_events = sync.get_events_detailed(store, target_cal)
native_canceled = 0
for ev in google_events:
    if sync.SYNC_TAG in ev["notes"]:
        continue  # Skip synced events, handled above
    is_canceled = ev["title"].lower().startswith("canceled:") or \
                  ev["title"].lower().startswith("cancelled:")
    if is_canceled:
        if store.removeEvent_span_error_(ev["event_ref"], EKSpanThisEvent, None):
            native_canceled += 1
            sync.log.info(f"  Removed: {ev['title']}")

sync.log.info(f"  Canceled native Google events removed: {native_canceled}")

sync.save_state(state)
sync.log.info("=== Done ===")
