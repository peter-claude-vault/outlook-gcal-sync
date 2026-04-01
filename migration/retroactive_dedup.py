#!/usr/bin/env python3
"""One-shot retroactive dedup for the past 14 days."""
import sync

# Override the lookback window
sync.SYNC_DAYS_BACK = 14
sync.SYNC_DAYS_FORWARD = 0  # Only past events

store = sync.get_event_store()
config = sync.load_config()
source_cal = sync.find_calendar(store, config["source_calendar"])
target_cal = sync.find_calendar(store, config["target_calendar"])

state = sync.load_state()
exclusions = sync.load_exclusions()

sync.log.info("Retroactive dedup: past 14 days...")
state, exclusions = sync.dedup(store, source_cal, target_cal, state, exclusions)

sync.save_state(state)
sync.save_exclusions(exclusions)
sync.log.info("Done.")
