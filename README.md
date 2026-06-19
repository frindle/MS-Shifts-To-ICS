# MS Shifts To ICS

Browser extensions that scrape your Microsoft Teams Shifts schedule and export it to a calendar (ICS / Outlook / iCloud).

## Download

### Firefox
> **[Download latest Firefox extension (.xpi)](https://github.com/frindle/MS-Shifts-To-ICS/releases/latest/download/teams-shifts-exporter-firefox-signed.xpi)**

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

### v1.61
- Fixed sign-up requests not being detected — tenant ID is now read directly from captured request URLs instead of relying on JWT decoding

### v1.60
- Renamed toggle to "Sync sign-ups to Open Time Signup calendar" to make it clear sign-ups go to a separate calendar

### v1.59
- Availability sign-up slots now sync to a separate **"Open Time Signup"** iCloud calendar instead of "Work Shifts"

### v1.58
- Added "Sync sign-ups to Open Time Signup calendar" toggle (default off) — syncs availability slots you've signed up for to a separate iCloud calendar
- Changed "Include open shifts" toggle to default off (was previously on by default)

### v1.57
- Availability sign-up slots you've signed up for now appear on your calendar with a "Signed Up: [shift]" prefix — previously all bare time-code open shifts were filtered out regardless of sign-up status

### v1.56
- Fixed future shifts missing from ICS/iCloud — API responses are paginated via `nextLink`; previous code only read the first page (which was full of availability sign-up shifts), so scheduled shifts on page 2+ were silently dropped

### v1.55
- Exclude availability sign-up open shifts (titles that are bare time codes like "1430 DX" or "0015 DX/DXC") — real posted open shifts with position prefixes like "P2 1245" are still included

### v1.54
- Added open shift debug data panel in popup (▸ Open shift debug data) to inspect raw API fields after sync — needed to identify how new availability sign-up shifts differ from regular open shifts

### v1.53
- Fixed "Download ICS" button throwing `TypeError: Cannot read properties of undefined (reading 'download')` — `chrome.downloads` is not available in offscreen documents; switched to anchor-click on blob URL

### v1.52
- Fixed Firefox cancel sync button causing popup to close
- Fixed Outlook calendar auto-focus unfocusing the Work Shifts calendar itself
- Fixed Chrome/Docker iCloud upload rate limiting (added missing delays)
- Improved error logging in background scripts for better debugging
- Error messages no longer persist across browser restarts (stale errors auto-clear)
- Fixed potential crash when stored history data is corrupted (now validates and resets gracefully)
- Improved "No teams found" error to suggest checking if schedule is set up in Teams
- General stability improvements

### v1.26
- Added TDO (Time off) support: TDO blocks are now exported as all-day calendar events with any notes included in the description
- Fixed shift notes containing newlines (e.g. traded shifts with multi-line trade notes) being silently dropped due to regex not matching across line breaks — notes now appear in the event description
- Regex hardened to handle overnight shifts where Teams includes a weekday prefix before the end date, and non-standard theme names

### v1.25.6
- Fixed sync failure when Shifts is not pinned to the left sidebar: flyout now waits up to 2s for the menu to render, uses broader item matching, and reports a clear error instead of silently timing out
- Navigation errors (sidebar not found, Shifts not in menu) now surface immediately with a descriptive message

### v1.25.5
- Added 90-second watchdog: if no progress is made for 90 seconds the sync is automatically cancelled and an error is reported, preventing silent hangs

### v1.25.4
- Fixed sync failure after "Almost there!" Continue click: wait for Teams sidebar to become interactive before re-navigating to Shifts (prevents clicking into an unresponsive page after the post-auth reload)

### v1.25.3
- Increased per-week stability threshold from 1.2s to 2s so shifts finish rendering before moving to the next week

### v1.25.2
- Chrome: port iCloud putEvent retry logic (was Firefox-only)

### v1.25.1
- Replace fixed 500ms per-week pause with a smart wait (up to 1500ms) that exits as soon as shift cards appear

### v1.25
- Increased week-load timeouts: shifts-stable wait raised to 10s max, week-change wait to 8s (both still exit early as soon as content is ready)

### v1.24.6 (Firefox)
- Fixed iCloud upload hanging mid-way by making the background page persistent (Firefox was suspending it during long syncs, pausing all fetches and timeouts)

### v1.24.5 (Firefox)
- Close the Teams scrape window immediately after scraping, before iCloud/Outlook sync runs
- Increased iCloud upload delay to 500ms to further reduce rate-limit stalls

### v1.24.4
- Fixed wrong "more apps" button being clicked (was selecting chat's "..." instead of the left sidebar one), causing "Timed out waiting for Shifts iframe"

### v1.24.3
- Firefox: add 250ms delay between iCloud uploads to prevent rate-limit stalls mid-upload

### v1.24.3 (Chrome) / v1.24.2 (Firefox)
- Firefox: retry iCloud uploads up to 3 times on stall/timeout to fix hang mid-upload
- Firefox: increased popup width by 5% to reduce scrollbars

### v1.24.3 (Chrome)
- Close the Teams tab opened for scraping after sync completes

### v1.24.1 (Firefox) / v1.24.2 (Chrome)
- Chrome: reduced post-Continue pause to 2500ms

### v1.24.1 (Firefox) / v1.23.1 (Chrome)
- Fixed sync getting stuck after clicking Continue on the "Almost there!" dialog (added delay to let Teams finish its post-auth reload before re-navigating)

### v1.24 (Firefox) / v1.23 (Chrome)
- Auto-dismiss Teams "Almost there!" permissions dialog when opening Shifts
- Fixed sync failure when Teams reloads the page after accepting the permissions dialog

### v1.22 / v1.22.1 (Firefox)
- Fixed long status messages causing horizontal scrollbar in popup
- Firefox: fixed double-scrape when sync was triggered concurrently

### v1.21
- Firefox: fixed ICS download blocked by browser security (switched to blob URL)

### v1.20
- Chrome: fixed Outlook import window minimizing macOS windows (now opens as background tab)

### v1.19
- iCloud app password now persists between extension updates
- Chrome: fixed ICS download (was blocked by MV3 service worker restrictions)
- Chrome: fixed "stuck on Opening Teams" when Teams was already open
- Firefox: fixed week X of XX progress during scraping
- Firefox: sync now runs in a background window
- Show last sync error in popup if previous sync failed
- Update banner in popup when a newer version is available

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
