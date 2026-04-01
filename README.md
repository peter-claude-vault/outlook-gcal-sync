# Outlook → Google Calendar Sync

See all your L'Oréal meetings on your Google Calendar — automatically.

If you're an Artefact consultant on L'Oréal, you have two calendars: an Exchange calendar (L'Oréal) and a Google calendar (Artefact). This tool copies your L'Oréal meetings to your Google Calendar every 10 minutes, so you only need to check one place. It also handles the annoying case where you're invited to the same meeting on both calendars — it keeps one copy and removes the duplicate.

**No API keys. No IT approvals. No cloud setup.** It uses the calendar accounts you've already added on your Mac.

---

## Step 1: Set Up Your Calendar Accounts (Do This First)

Before installing anything, you need both your L'Oréal and Artefact calendar accounts connected to your Mac. This is the most important step — **if this isn't done correctly, nothing else will work.**

### Add your L'Oréal Exchange account

1. Open **System Settings** on your Mac (click the Apple menu  in the top-left corner of your screen → System Settings)
2. Click **Internet Accounts** in the left sidebar
3. Look for your L'Oréal account (it will say "Exchange" and show your `@loreal.com` email)
4. **If it's already there:** Click on it and make sure **Calendars** is toggled ON
5. **If it's NOT there:** Click the **Add Account...** button → choose **Microsoft Exchange** → sign in with your L'Oréal email and password

### Add your Artefact Google account

1. Still in **System Settings → Internet Accounts**
2. Look for your Artefact account (it will say "Google" and show your `@artefact.com` email)
3. **If it's already there:** Click on it and make sure **Calendars** is toggled ON
4. **If it's NOT there:** Click the **Add Account...** button → choose **Google** → sign in with your Artefact email

### Verify both calendars are working

1. Open the **Calendar** app on your Mac (search for "Calendar" in Spotlight, or find it in your Applications folder)
2. You should see events from **both** your L'Oréal and Artefact calendars
3. If you only see one calendar's events, go back to System Settings and check that Calendars is enabled for both accounts

> **Stop here if you don't see both calendars in the Calendar app.** The sync tool reads from these calendars — if they're not showing up in Apple Calendar, the tool can't access them either.

---

## Step 2: Install Python (One-Time Setup)

The sync tool is written in Python. Most Macs don't have it pre-installed, so let's check.

### Open Terminal

Terminal is an app on your Mac that lets you type commands. Here's how to open it:

1. Press **Command (⌘) + Space** on your keyboard — this opens Spotlight search
2. Type **Terminal**
3. Press **Enter**

A window will open with a dark or light background and a blinking cursor. This is where you'll type (or paste) all the commands in this guide. **Keep this window open** — you'll use it for the rest of the setup.

### Check if Python is already installed

Copy the line below, paste it into Terminal (Command + V), and press **Enter**:

```
python3 --version
```

- **If you see something like `Python 3.12.4`** — great, Python is installed. Skip ahead to **Step 3**.
- **If you see `command not found` or any error** — follow the steps below to install it.

### Install Homebrew (a tool that installs other tools)

Copy and paste this entire line into Terminal and press **Enter**:

```
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
```

- It will ask for your Mac password. **When you type your password, nothing will appear on screen** — that's normal, it's a security feature. Just type it and press Enter.
- This takes a few minutes. Wait for it to finish before moving on.

### Install Python

Copy and paste this into Terminal and press **Enter**:

```
brew install python
```

Wait for it to finish, then verify it worked:

```
python3 --version
```

You should now see a version number like `Python 3.12.4`.

---

## Step 3: Install the Sync Tool

You should still have Terminal open from the previous step. If you closed it, open it again (Command + Space → type "Terminal" → Enter).

Copy and paste these commands **one at a time**. After pasting each one, press **Enter** and wait for it to finish before pasting the next one.

**Command 1 — Download the tool:**
```
git clone https://github.com/peter-claude-vault/outlook-gcal-sync.git ~/outlook-gcal-sync
```
This downloads the tool to a folder called `outlook-gcal-sync` in your home directory.

**Command 2 — Go into the folder:**
```
cd ~/outlook-gcal-sync
```

**Command 3 — Run the installer:**
```
./install.sh
```

### What happens next

The installer will walk you through a few prompts:

1. **It installs dependencies** — this happens automatically, no input needed from you. Wait for it to finish.

2. **It asks you to pick your L'Oréal calendar.** You'll see a numbered list of all calendars on your Mac, like this:
   ```
   [0] Calendar  (source: Exchange)  ← Exchange
   [1] your.name@artefact.com  (source: Google)  ← Google
   ```
   Type the **number** next to your L'Oréal Exchange calendar and press **Enter**. (It's usually called "Calendar" with an "Exchange" source.)

3. **It asks you to pick your Google calendar.** Type the **number** next to your Artefact Google calendar (it will show your `@artefact.com` email) and press **Enter**.

4. **It runs an initial sync.** You'll see output showing how many events were created. Something like:
   ```
   Done. Created=45 Updated=0 Deleted=0 Unchanged=0 Skipped=12
   ```

5. **It sets up automatic syncing.** From this point on, your Mac will sync every 10 minutes in the background. This keeps working even after you restart your computer. **You never need to open Terminal again.**

### Verify it worked

Open **Google Calendar** in your web browser. You should see your L'Oréal meetings appearing alongside your Artefact meetings. Synced events will have `[outlook-sync]` in their notes if you click on them.

---

## What Happens With Duplicate Meetings

When you're invited to the same meeting on both calendars (this happens a lot — someone sends a Teams invite that goes to both your L'Oréal and Artefact emails), the tool automatically keeps one copy and removes the other:

- **Meeting organized by someone at Artefact** → keeps the Google Calendar version
- **Meeting organized by anyone else** (L'Oréal, Sapient, external vendors) → keeps the Exchange version

You don't need to do anything. Duplicates are resolved automatically every 10 minutes.

---

## Troubleshooting

### "Calendar access denied"

Your Mac is blocking the sync tool from reading your calendars. To fix this:

1. Open **System Settings** (Apple menu → System Settings)
2. Click **Privacy & Security** in the left sidebar
3. Scroll down and click **Calendars**
4. Find **Terminal** in the list and make sure the toggle is **ON**
5. If Terminal isn't in the list, try running the installer again — open Terminal and paste:
   ```
   cd ~/outlook-gcal-sync && ./install.sh
   ```
   Your Mac should prompt you to grant calendar access.

### "Calendar 'X' not found"

The calendar name saved during setup doesn't match what's on your Mac anymore. To fix this, re-run the setup:

1. Open Terminal (Command + Space → type "Terminal" → Enter)
2. Paste this and press Enter:
   ```
   cd ~/outlook-gcal-sync && .venv/bin/python3 sync.py --setup
   ```
3. Pick the correct calendars from the numbered list

### Events aren't showing up on Google Calendar

A few things to check:

1. **Check Apple Calendar first.** Open the Calendar app on your Mac. Do you see your L'Oréal events there? If not, the problem is your Exchange account setup (go back to Step 1), not this tool.

2. **Check the sync log.** Open Terminal and paste:
   ```
   cat ~/outlook-gcal-sync/sync.log
   ```
   If you see lines like `Created: Meeting Name` or `Unchanged: 139`, the tool is working correctly. If you see errors, something is wrong — reach out for help.

3. **Wait 10 minutes.** The sync runs on a 10-minute cycle. If you just installed it, give it one cycle to catch up.

### I still see duplicate events

The dedup process runs every 10 minutes. If you just installed the tool, give it 20–30 minutes to clean up existing duplicates. If they persist after that, check the sync log for errors.

---

## Pausing, Resuming, or Uninstalling

Open Terminal (Command + Space → type "Terminal" → Enter), then paste the relevant command:

**Pause syncing** (stops the 10-minute cycle, keeps everything installed):
```
launchctl unload ~/Library/LaunchAgents/com.outlook-gcal-sync.plist
```

**Resume syncing** (restarts the 10-minute cycle):
```
launchctl load ~/Library/LaunchAgents/com.outlook-gcal-sync.plist
```

**Completely uninstall** (removes everything):
```
launchctl unload ~/Library/LaunchAgents/com.outlook-gcal-sync.plist
rm ~/Library/LaunchAgents/com.outlook-gcal-sync.plist
rm -rf ~/outlook-gcal-sync
```

Events that were already synced to Google Calendar will stay — uninstalling only stops future syncing.

---

## Migrating from OGCS

If you previously used [Outlook Google Calendar Sync (OGCS)](https://github.com/phw198/OutlookGoogleCalendarSync), you should clean up its leftover events before starting. See [migration/README.md](migration/README.md) for instructions.
