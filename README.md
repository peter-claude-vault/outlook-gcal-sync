# Outlook → Google Calendar Sync

See all your L'Oréal meetings on your Google Calendar — automatically.

If you're an Artefact consultant on L'Oréal, you have two calendars: an Exchange calendar (L'Oréal) and a Google calendar (Artefact). This tool copies your L'Oréal meetings to your Google Calendar every 10 minutes, so you only need to check one place. It also handles the annoying case where you're invited to the same meeting on both calendars — it keeps one copy and removes the duplicate.

**No API keys. No IT approvals. No cloud setup.** It uses the calendar accounts you've already added on your Mac.

---

## Before You Start

You need three things:

1. **A Mac.** This only works on macOS. (It uses Apple's built-in calendar system.)

2. **Both calendar accounts added to your Mac.** Open **System Settings → Internet Accounts** and make sure you see:
   - Your L'Oréal Exchange account (the one with your `@loreal.com` email)
   - Your Artefact Google account (the one with your `@artefact.com` email)

   If either is missing, click the **+** button and add it. You should be able to see events from both accounts in Apple's Calendar app before proceeding.

3. **Python installed.** Open Terminal (search for "Terminal" in Spotlight) and type:
   ```
   python3 --version
   ```
   If you see a version number (3.10 or higher), you're good. If you get an error, install it:
   ```
   brew install python
   ```
   If `brew` doesn't work either, install Homebrew first by pasting this into Terminal:
   ```
   /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
   ```

---

## Installation

Open Terminal and paste these three lines:

```bash
git clone https://github.com/peter-claude-vault/outlook-gcal-sync.git ~/outlook-gcal-sync
cd ~/outlook-gcal-sync
./install.sh
```

The installer will walk you through everything:

1. It installs the Python dependencies (automatically)
2. It shows you a list of calendars on your Mac — **pick your L'Oréal calendar** (the Exchange one), then **pick your Google calendar** (the Artefact one)
3. It runs an initial sync so you can verify it works
4. It sets up automatic syncing every 10 minutes — you don't need to do anything after this

**That's it.** Your L'Oréal meetings will now appear on your Google Calendar. Open Google Calendar and check.

---

## What Happens With Duplicate Meetings

When you're invited to the same meeting on both calendars (this is common — someone sends a Teams invite that goes to both your L'Oréal and Artefact emails), the tool automatically picks one and removes the other:

- **Meeting organized by someone at Artefact** → keeps the Google Calendar version
- **Meeting organized by anyone else** (L'Oréal, Sapient, external) → keeps the Exchange version

This means you'll never see the same meeting twice.

---

## How to Check If It's Working

- **Logs:** Open Terminal and run:
  ```
  cat ~/outlook-gcal-sync/sync.log
  ```
  You should see entries like `Created: Meeting Name` or `Unchanged: X` every 10 minutes.

- **Google Calendar:** L'Oréal meetings should appear. Synced events have `[outlook-sync]` in their notes field.

---

## Troubleshooting

### "Calendar access denied"
Your Mac is blocking the script from reading calendars. Fix it:
1. Open **System Settings → Privacy & Security → Calendars**
2. Find **Terminal** in the list and make sure it's enabled
3. If Terminal isn't listed, try running the install again — it should prompt you

### "Calendar 'X' not found"
The calendar name in your config doesn't match what's on your Mac. Run:
```
cd ~/outlook-gcal-sync
.venv/bin/python3 sync.py --list-calendars
```
This shows all available calendars. Then re-run setup:
```
.venv/bin/python3 sync.py --setup
```

### Events not appearing on Google Calendar
- Open Apple's **Calendar** app and check that L'Oréal events actually show up there. If they don't, the problem is your Exchange account setup, not this tool.
- Make sure your Google account is added in **System Settings → Internet Accounts** (not just in Google Calendar's web interface).
- Check the log: `cat ~/outlook-gcal-sync/sync.log`

### I see duplicate events
The dedup runs every 10 minutes. If you just installed, give it 1-2 cycles (10-20 minutes) to clean up existing duplicates. If duplicates persist after 30 minutes, check the log for errors.

---

## Stopping or Uninstalling

**To pause syncing:**
```
launchctl unload ~/Library/LaunchAgents/com.outlook-gcal-sync.plist
```

**To resume syncing:**
```
launchctl load ~/Library/LaunchAgents/com.outlook-gcal-sync.plist
```

**To fully uninstall:**
```
launchctl unload ~/Library/LaunchAgents/com.outlook-gcal-sync.plist
rm ~/Library/LaunchAgents/com.outlook-gcal-sync.plist
rm -rf ~/outlook-gcal-sync
```

Events that were already synced to Google Calendar will remain — the uninstall only stops future syncing.

---

## Migrating from OGCS

If you previously used [Outlook Google Calendar Sync (OGCS)](https://github.com/phw198/OutlookGoogleCalendarSync), you should clean up its leftover events before starting. See [migration/README.md](migration/README.md) for instructions.
