# Migration from OGCS

If you're replacing [Outlook Google Calendar Sync (OGCS)](https://github.com/phw198/OutlookGoogleCalendarSync)
with this tool, run the OGCS cleanup **once before your first sync**.

OGCS creates duplicate calendar events with `.ogcs` markers in attendee names
and organizer fields. The standard sync dedup rules won't catch these because
they don't follow normal organizer-domain patterns.

## Steps

1. Complete the main setup first (`install.sh` or `sync.py --setup`)
2. Stop OGCS (uninstall or disable it)
3. Preview what will be cleaned up:
   ```bash
   .venv/bin/python3 migration/cleanup_ogcs.py
   ```
4. If the preview looks right, delete the OGCS events:
   ```bash
   .venv/bin/python3 migration/cleanup_ogcs.py --delete
   ```
5. Run the sync normally — it will create clean copies of Exchange events

The cleanup scans your Google calendar (next 30 days) for events with `.ogcs`
in their organizer, notes, or attendee fields, and removes them via EventKit.
