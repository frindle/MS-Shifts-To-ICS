# MS Shifts To ICS

Browser extensions that scrape your Microsoft Teams Shifts schedule and export it to a calendar (ICS / Outlook / iCloud).

## Download

### Firefox
> **[Download latest Firefox extension (.xpi)](https://github.com/frindle/MS-Shifts-To-ICS/releases/latest/download/teams-shifts-exporter-firefox.xpi)**

1. Click the link above to download the `.xpi` file
2. In Firefox, go to `about:addons` → gear icon → **Install Add-on From File…**
3. Select the downloaded `.xpi`

### Chrome / Edge
Load unpacked from the `teams-shifts-exporter-chrome` folder:

1. Go to `chrome://extensions` and enable **Developer mode**
2. Click **Load unpacked** and select the `teams-shifts-exporter-chrome` folder from this repo

## Features

- Scrapes your Teams Shifts schedule (including open shifts)
- Exports to ICS file, Outlook Calendar, or iCloud Calendar
- Auto-syncs on a daily schedule
- Skips open shifts that overlap or are within 8 hours of your assigned shifts
- Supports midnight-crossing shifts

## Usage

1. Open Microsoft Teams Shifts in your browser
2. Click the extension icon and hit **Sync Shifts**
3. Optionally enable **Outlook Calendar** or **iCloud Calendar** sync in the popup

## Changelog

### v1.17
- Open shift 8-hour gap filter is now always applied (no longer optional)
- Added calendar icon
- Firefox: fixed stale sync state on browser restart

### v1.16
- Added cancel sync button
- Fixed progress bar during iCloud upload phase
- Fixed "removing old shifts" step getting stuck
- Fixed sync button stuck disabled after Chrome restart mid-sync

### v1.15
- Added progress bar with step labels
- Fixed midnight-crossing open shifts (e.g. "Apr 28, 12:30 AM" end times)
- Fixed UID collision between open shifts and scheduled shifts with same start time
- Added iCloud open shift tracking to prevent re-upload issues
