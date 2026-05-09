#!/bin/bash
set -e

# Start a virtual framebuffer (Chrome requires a display even in Docker)
Xvfb :99 -screen 0 1280x900x24 -nolisten tcp &
export DISPLAY=:99
sleep 1

# Start VNC server — connect on port 5900 for the initial Teams login
x11vnc -display :99 -forever -nopw -quiet &

echo "[entrypoint] Container ready. VNC on port 5900 for initial Teams login."
echo "[entrypoint] Running first sync..."

# Run immediately on startup
node /app/runner/sync.js || true

# Then repeat every 24 hours
while true; do
  echo "[entrypoint] Sleeping 24 hours..."
  sleep 86400
  echo "[entrypoint] Running scheduled sync..."
  node /app/runner/sync.js || true
done
