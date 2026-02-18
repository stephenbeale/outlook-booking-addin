# Outlook Room Booking Add-in

An Outlook Web Add-in that opens a pre-filled compose window for sending weekly room booking requests to kingsresidences@kcl.ac.uk. The email body automatically includes a date 28 days in the future.

## Prerequisites

- Node.js (for the local HTTPS server)
- Outlook desktop or Outlook on the web (outlook.office.com)

## Setup

### 1. Install dev certificates

Generate trusted self-signed certificates so Office can load the add-in over HTTPS:

```bash
npx office-addin-dev-certs install
```

This creates certificates in your home directory. Note the paths it prints.

### 2. Serve the add-in over HTTPS

From the project folder:

```bash
npx http-server -S -C ~/.office-addin-dev-certs/localhost.crt -K ~/.office-addin-dev-certs/localhost.key -p 3000
```

Or with the shorthand (if the certs are in the default location):

```bash
npx http-server -S -p 3000
```

The taskpane will be available at `https://localhost:3000/taskpane.html`.

### 3. Sideload the manifest

**Outlook on the web:**
1. Go to https://outlook.office.com
2. Click the gear icon > **Get Add-ins** (or **Manage Add-ins**)
3. Click **My add-ins** > **Add a custom add-in** > **Add from file...**
4. Select `manifest.xml` from this folder

**Outlook desktop (Windows):**
1. Open Outlook
2. Go to **File** > **Manage Add-ins** (opens the web interface)
3. Follow the same steps as Outlook on the web above

**Alternative (Windows shared folder):**
1. Create a shared network folder or use `\\localhost\c$\path\to\outlook-booking-addin`
2. In Outlook: **File** > **Options** > **Trust Center** > **Trust Center Settings** > **Trusted Add-in Catalogs**
3. Add the folder path as a catalog URL
4. Restart Outlook, then find the add-in under **My Add-ins** > **Shared Folder**

## Usage

1. Open any email in Outlook (or start a new message)
2. Click the **Book Room** button in the ribbon
3. The task pane opens showing the booking date (28 days from today)
4. Click **Compose Booking Email**
5. A new compose window opens with To, Subject, and Body pre-filled
6. Review and click Send

## Project Structure

```
manifest.xml     - Office Add-in manifest (sideload config)
taskpane.html    - Task pane UI
taskpane.js      - Date calculation + Office.js compose logic
taskpane.css     - Styling
README.md        - This file
```
