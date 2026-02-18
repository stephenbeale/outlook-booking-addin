# CLAUDE.md — Outlook Room Booking Add-in

## Project Overview

An Outlook Web Add-in that opens a pre-filled compose window for sending weekly room booking requests to kingsresidences@kcl.ac.uk. The email body automatically includes a date 28 days in the future.

**GitHub:** https://github.com/stephenbeale/outlook-booking-addin

## Key Details

- **Recipient:** kingsresidences@kcl.ac.uk
- **User account:** KCL (King's College London) Outlook account
- **Booking details hardcoded in taskpane.js:**
  - Full name: Stephen Beale
  - Residences: Great Dover Street
  - Room: Single
  - Duration: 1 night
  - Mobile: 07803571590
- **Date logic:** Calculates today + 28 days, formatted as e.g. "Wednesday 18 March"

## Project Structure

```
manifest.xml     - Office Add-in manifest (sideload config, points to https://localhost:3000)
taskpane.html    - Task pane UI
taskpane.js      - Date calculation + Office.js compose logic
taskpane.css     - Styling
README.md        - Setup and usage instructions
ROADMAP.md       - Future feature plans
```

## Development Workflow

### Starting the local HTTPS server

The manifest points to `https://localhost:3000`. You must run this server every session before the add-in will work:

```bash
cd C:\Users\sjbeale\source\outlook-booking-addin
npx http-server -S -C ~/.office-addin-dev-certs/localhost.crt -K ~/.office-addin-dev-certs/localhost.key -p 3000
```

Dev certificates were already installed via `npx office-addin-dev-certs install` — no need to reinstall.

Verify the server is working by visiting https://localhost:3000/taskpane.html in a browser (accept the self-signed cert warning if prompted).

### Sideloading the manifest into Outlook

**Outlook on the web (primary method for KCL account):**
1. Go to https://outlook.office.com (sign in with KCL account)
2. Click the gear icon (Settings) > **Get Add-ins** (or **Manage Add-ins**)
3. Click **My add-ins** > **Add a custom add-in** > **Add from file...**
4. Select `manifest.xml` from `C:\Users\sjbeale\source\outlook-booking-addin\`
5. The "Book Room" button should appear in the ribbon when reading a message

## Session Notes

### 2026-02-18 - Session Summary

**Work Completed:**
- Created the full Outlook Web Add-in project from scratch with all 5 source files
- Implemented date calculation (today + 28 days) with human-readable formatting
- Implemented `composeEmail()` using `Office.context.mailbox.displayNewMessageForm()` to open a pre-filled compose window
- Included HTML email body with bullet-point booking details and a plain-text fallback via mailto for testing outside Outlook
- Initialized git repo, used feature branch (`feature/outlook-booking-addin`), merged to main
- Pushed all commits to GitHub: https://github.com/stephenbeale/outlook-booking-addin
- Added ROADMAP.md with future plans (recurring reminders, dynamic fields, etc.)
- Installed `office-addin-dev-certs` for trusted HTTPS (`npx office-addin-dev-certs install`)
- Started local HTTPS server on port 3000 (was running as background task — will need restarting next session)

**Work In Progress:**
- Sideloading `manifest.xml` into Outlook on the web — this is the immediate next step

**Unfinished Git Workflows:**
- None. All work is committed and pushed. Working tree is clean. No open PRs.

**Next Steps:**
1. Re-start the HTTPS server (see command above)
2. Verify https://localhost:3000/taskpane.html loads in the browser
3. Sideload `manifest.xml` into Outlook on the web (see sideloading steps above)
4. Test the add-in: open any email, click "Book Room" in the ribbon, verify compose window opens with correct pre-filled content
5. If the manifest fails to load, check that the HTTPS server is running and the cert is trusted

**Technical Notes:**
- The manifest uses a placeholder GUID (`a1b2c3d4-e5f6-7890-abcd-ef1234567890`) — fine for sideloading, would need a real UUID for AppSource publishing
- The manifest references icon files (`icon-16.png`, `icon-32.png`, `icon-80.png`) that do not yet exist in the repo. Outlook may show a broken icon but the add-in should still function. Add placeholder PNGs if icon errors cause manifest rejection.
- KCL may have MDM/IT policies that restrict custom add-in sideloading. If sideloading is blocked, contact KCL IT or try the Outlook desktop shared folder method (documented in README.md)
- The `office-addin-dev-certs` package installs certs to `~/.office-addin-dev-certs/` (i.e. `C:\Users\sjbeale\.office-addin-dev-certs\`)
