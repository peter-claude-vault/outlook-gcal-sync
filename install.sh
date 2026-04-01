#!/bin/bash
#
# Outlook → Google Calendar Sync — Installer
#
# Sets up Python venv, installs dependencies, runs the setup wizard,
# generates and loads the launchd plist for automatic 10-minute syncing.
#
set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PLIST_NAME="com.outlook-gcal-sync"
PLIST_PATH="$HOME/Library/LaunchAgents/$PLIST_NAME.plist"

echo "=== Outlook → Google Calendar Sync Installer ==="
echo ""

# 1. Check Python
if ! command -v python3 &>/dev/null; then
    echo "Error: python3 not found. Install via: brew install python"
    exit 1
fi

PYTHON_VERSION=$(python3 -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')")
echo "Found Python $PYTHON_VERSION"

# 2. Create venv and install deps
if [ ! -d "$SCRIPT_DIR/.venv" ]; then
    echo "Creating virtual environment..."
    python3 -m venv "$SCRIPT_DIR/.venv"
fi

echo "Installing dependencies..."
"$SCRIPT_DIR/.venv/bin/pip" install -q -r "$SCRIPT_DIR/requirements.txt"

# 3. Run setup wizard (picks source/target calendars)
echo ""
"$SCRIPT_DIR/.venv/bin/python3" "$SCRIPT_DIR/sync.py" --setup

# 4. Test sync
echo ""
echo "Running initial sync..."
"$SCRIPT_DIR/.venv/bin/python3" "$SCRIPT_DIR/sync.py"

# 5. Generate and load launchd plist
echo ""
echo "Setting up automatic sync (every 10 minutes)..."

cat > "$PLIST_PATH" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>$PLIST_NAME</string>
    <key>ProgramArguments</key>
    <array>
        <string>$SCRIPT_DIR/.venv/bin/python3</string>
        <string>$SCRIPT_DIR/sync.py</string>
    </array>
    <key>StartInterval</key>
    <integer>600</integer>
    <key>StandardOutPath</key>
    <string>$SCRIPT_DIR/sync.log</string>
    <key>StandardErrorPath</key>
    <string>$SCRIPT_DIR/sync.log</string>
    <key>RunAtLoad</key>
    <true/>
</dict>
</plist>
EOF

launchctl unload "$PLIST_PATH" 2>/dev/null || true
launchctl load "$PLIST_PATH"

echo ""
echo "=== Done ==="
echo "Sync is running every 10 minutes."
echo "Logs: $SCRIPT_DIR/sync.log"
echo ""
echo "To stop:   launchctl unload $PLIST_PATH"
echo "To restart: launchctl load $PLIST_PATH"
