# Interlink 2026 Registration Portal

A modern, responsive, and glassmorphic registration portal built with **Bootstrap 5**, **Vanilla JavaScript**, and a **Google Apps Script** backend.

---

## 📁 Repository Structure

*   **[index.html](file:///wsl.localhost/Ubuntu/opt/interlink/index.html)**: The frontend page containing the HTML layout, responsive Bootstrap grid structure, custom styling rules (including keyframe animations and theme colors), and inline script config.
*   **[app.js](file:///wsl.localhost/Ubuntu/opt/interlink/app.js)**: The modular JavaScript logic containing the validation logic, conditional visibility toggle handlers, and the fetch submission configuration.
*   **[google-apps-script.js](file:///wsl.localhost/Ubuntu/opt/interlink/google-apps-script.js)**: The backend code to be copied and deployed as a Google Apps Script Web App.

---

## ✨ Features

1.  **Premium Glassmorphic Design**: Clean UI featuring a dark theme, shifting ambient color blobs, glowing input focus states, and gradient title branding.
2.  **Modular Frontend Separation**: Keeps HTML and JS separated for neat code maintenance (with inline script fallback).
3.  **Vanilla JS Conditional Logic**: Toggles school program selectors and custom specify-text inputs dynamically with smooth height transitions:
    *   *CSPC selected*: Shows CSPC-specific program dropdown. Hides UNAIR. Removes `specifyProgram` options.
    *   *UNAIR selected*: Shows UNAIR-specific program dropdown. Hides CSPC.
    *   *UNAIR "Other" selected*: Dynamically exposes the `specifyProgram` text input.
    *   *Year Level "Other" selected*: Exposes the `specifyYearLevel` text input.
4.  **Auto-Clearing Inactive Fields**: When fields are hidden by the visibility logic, their values are reset to empty strings (`""`) so that incorrect program entries are never written to the sheet.
5.  **Interactive Submission States**: Integrates a full-screen blur overlay containing a modern spinning wheel, a green SVG success checkmark, or a red failure cross depending on the network response.
6.  **Auto-Initializing Backend Sheet**: The Google Apps Script automatically checks if the spreadsheet tab `"Interlink2026"` exists, creating it dynamically with formatted headers and auto-freeze rules if absent.
7.  **Email Duplicate Interception**: Google Apps Script cross-references emails against existing row inputs to prevent users from registering multiple times.
8.  **Personalized Confirmation Emails**: Automatically constructs and sends professional HTML emails tailored specifically to either CSPC or UNAIR registrants, aligned with the event's slate color design.
9.  **Google Drive Virtual Background Attachment**: Attaches a custom MS Teams virtual background downloaded directly from Google Drive during registration.
10. **High-Concurrency Protection (Thread Locking)**: Uses Apps Script `LockService` to serialize database access, preventing data overlaps or write failures when multiple users submit registrations at the exact same moment.
11. **Performance-Optimized Execution**: The script releases the thread lock *immediately* after saving the registrant's data to the spreadsheet. Slow operations, like compiling and sending emails, are performed outside the locked block to optimize system speed.

---

## 🚀 Deployment Instructions

### Step 1: Set Up the Google Spreadsheet
1.  Open your Google Drive and create a new **Google Sheet** (or open an existing one).
2.  Name the spreadsheet (e.g. `Interlink 2026 Registrations`).

### Step 2: Paste the Apps Script Backend
1.  In your Google Sheet menu bar, click on **Extensions** > **Apps Script**.
2.  Delete any default template code in the editor (usually `function myFunction() {}`).
3.  Open **[google-apps-script.js](file:///wsl.localhost/Ubuntu/opt/interlink/google-apps-script.js)**, copy the entire file contents, and paste it into the Apps Script editor.
4.  Configure the constants at the top of the script:
    *   `VIRTUAL_BACKGROUND_DRIVE_ID`: Set the file ID of your virtual background image from Google Drive.
    *   `MsTeams_LINK`: Microsoft Teams link for the webinar.
    *   `EVENT_DATE`: E.g., `June 03, 2026`.
    *   `SENDER_NAME`: The display name for automated emails.
5.  Click the **Save (Disk icon)** button.

### Step 3: Deploy the Script as a Web App
1.  In the top-right corner of the Apps Script workspace, click the **Deploy** button and select **New deployment**.
2.  Click the **Select type (Gear icon)** and select **Web app**.
3.  Configure the settings:
    *   **Description**: `Interlink 2026 API`
    *   **Execute as**: `Me (your-email@gmail.com)`
    *   **Who has access**: `Anyone` *(Crucial: This allows your frontend to send requests to it)*.
4.  Click **Deploy**.
5.  If prompted, click **Authorize access**, log in to your Google Account, click **Advanced** > **Go to Untitled project (unsafe)**, and click **Allow**.
6.  Once deployed, copy the generated **Web app URL** (ends in `/exec`).

### Step 4: Link Frontend to the Web App
1. Open **[app.js](file:///wsl.localhost/Ubuntu/opt/interlink/app.js)**.
2. Locate the JavaScript configuration section at the top of the file:
   ```javascript
   const WEB_APP_URL = "YOUR_SCRIPT_URL_HERE";
   ```
3. Replace `"YOUR_SCRIPT_URL_HERE"` with the copied Google Apps Script Web App URL.
4. Open **[index.html](file:///wsl.localhost/Ubuntu/opt/interlink/index.html)**.
5. If using the inline script version, locate the JavaScript configuration section around line 620:
   ```javascript
   const ENDPOINT_URL = "YOUR_SCRIPT_URL_HERE";
   ```
6. Replace it with the same Web App URL.
7. Save the files and deploy to **Netlify**.

---

## ⚙️ REST API Key Mapping
The frontend outputs standard JSON objects payload to the Apps Script endpoint matching this schema:

| Key | Description | Optional |
|---|---|---|
| `firstName` | Student first name | No |
| `middleInitial` | Middle initial | Yes |
| `lastName` | Student last name | No |
| `suffix` | Generational suffix (e.g., Jr.) | Yes |
| `email` | Registered email address (checked for duplicates) | No |
| `school` | `"CSPC"` or `"UNAIR"` | No |
| `programCSPC` | `"BSIS"`, `"BSIT"`, `"BSCS"`, `"BLIS"` (or `""` if not CSPC) | Conditionally Required |
| `programUNAIR` | `"Informatic Engineering"`, `"Other"` (or `""` if not UNAIR) | Conditionally Required |
| `specifyProgram` | Custom text input (if UNAIR "Other" selected, else `""`) | Conditionally Required |
| `yearLevel` | `"1st Year"`, `"2nd Year"`, etc. | No |
| `specifyYearLevel` | Custom text input (if Year Level "Other" selected, else `""`) | Conditionally Required |
