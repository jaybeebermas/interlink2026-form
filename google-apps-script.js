/**
 * Deploys as a Google Apps Script Web App.
 * Intercepts POST requests from the frontend, validates fields, checks for duplicates,
 * appends registration data to the spreadsheet, and sends a personalized invitation email.
 */

// ============================================================
// CONFIGURATION — Edit these values before deploying
// ============================================================

// Google Drive File ID of the MS Teams virtual background image.
// 1. Upload the image to Google Drive.
// 2. Right-click → "Get link" → set access to "Anyone with the link".
// 3. Copy the File ID from the URL (the long string between /d/ and /view).
// Leave as "" to skip the attachment.
const VIRTUAL_BACKGROUND_DRIVE_ID = "1DovIID3oNuH7ilmnUftE6Z7x5NOhsjTq";

// Event details
const MsTeams_LINK = "https://teams.microsoft.com/meet/4189155944193?p=2b7O8JnKNZP5AWpH9U";
const MEETING_ID = "418 915 594 419 3";
const MEETING_PASSCODE = "Vu3FJ2Ga";
const EVENT_DATE = "June 03, 2026";
const CSPC_FB_URL = "https://www.facebook.com/cspc.ccs";
const SENDER_NAME = "Interlink 2026";

// ============================================================


/**
 * Entry point: Handles POST requests from the registration form.
 */
function doPost(e) {
  // Check if form is still accepting responses (Deadline: June 3, 2026, 8:30 AM UTC+8)
  const deadline = 1780446600000; // Unix timestamp for 2026-06-03T08:30:00+08:00
  if (new Date().getTime() >= deadline) {
    return createJsonResponse(false, "This form is no longer accepting responses. Registration closed on June 3, 2026, at 8:30 AM.");
  }

  const lock = LockService.getScriptLock();
  try {
    // Wait for up to 30 seconds for the lock to become available
    lock.waitLock(30000);
  } catch (lockError) {
    return createJsonResponse(false, "The server is currently busy. Please try again in a few seconds.");
  }

  let data;
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName("Interlink2026");

    // Dynamic Sheet Initialization
    if (!sheet) {
      sheet = spreadsheet.insertSheet("Interlink2026");
      sheet.appendRow([
        "First Name",
        "Middle Initial",
        "Last Name",
        "Suffix",
        "Email",
        "School",
        "Program (CSPC)",
        "Program (UNAIR)",
        "Specified Program",
        "Year Level",
        "Specified Year Level",
        "Registration Date"
      ]);
      const headerRange = sheet.getRange(1, 1, 1, 12);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#f1f5f9");
      headerRange.setFontColor("#0f172a");
      sheet.setFrozenRows(1);
    }

    // Parse JSON payload from POST request
    data = JSON.parse(e.postData.contents);

    // Backend validation
    const email = (data.email || "").trim().toLowerCase();
    const firstName = (data.firstName || "").trim();
    const lastName = (data.lastName || "").trim();
    const school = (data.school || "").trim();
    const yearLevel = (data.yearLevel || "").trim();

    if (!email || !firstName || !lastName || !school || !yearLevel) {
      return createJsonResponse(false, "Required fields are missing. Please fill in the entire form.");
    }

    // Duplicate check
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const existingEmail = rows[i][4].toString().trim().toLowerCase();
      if (existingEmail === email) {
        return createJsonResponse(false, `The email address "${data.email}" is already registered.`);
      }
    }

    // Write row to sheet
    sheet.appendRow([
      data.firstName || "",
      data.middleInitial || "",
      data.lastName || "",
      data.suffix || "",
      data.email || "",
      data.school || "",
      data.programCSPC || "",
      data.programUNAIR || "",
      data.specifyProgram || "",
      data.yearLevel || "",
      data.specifyYearLevel || "",
      new Date()
    ]);

    // Flush spreadsheet changes to disk before releasing the lock
    SpreadsheetApp.flush();

  } catch (error) {
    return createJsonResponse(false, "Server Error: " + error.toString());
  } finally {
    // Release the lock so other pending requests can run
    lock.releaseLock();
  }

  // Send invitation email (outside the lock block since email delivery is slow and doesn't modify the spreadsheet)
  if (data) {
    try {
      sendInvitationEmail(data);
    } catch (mailError) {
      // Log the error for debugging but do not fail the registration
      console.error("Email send failed for " + data.email + ": " + mailError.toString());
    }
  }

  return createJsonResponse(true, "Registration successful! Welcome to Interlink 2026.");
}


/**
 * Sends a personalized HTML invitation email to the registrant.
 * Automatically selects the CSPC or UNAIR template based on the school field.
 *
 * @param {Object} data - The parsed registration payload from the form.
 */
function sendInvitationEmail(data) {
  const recipientEmail = (data.email || "").trim();
  const firstName = (data.firstName || "").trim();
  const school = (data.school || "").trim().toUpperCase();

  if (!recipientEmail || !firstName || !school) return;

  // Build options object
  const mailOptions = {
    name: SENDER_NAME,
  };

  // Attach virtual background if a Drive ID is configured
  let bgBlob = null;
  if (VIRTUAL_BACKGROUND_DRIVE_ID && VIRTUAL_BACKGROUND_DRIVE_ID.trim() !== "") {
    try {
      const bgFile = DriveApp.getFileById(VIRTUAL_BACKGROUND_DRIVE_ID.trim());
      bgBlob = bgFile.getBlob();
      mailOptions.inlineImages = {
        virtual_bg: bgBlob
      };
    } catch (driveErr) {
      console.warn("Could not retrieve virtual background from Drive: " + driveErr.toString());
    }
  }

  let subject, htmlBody;

  if (school === "CSPC") {
    subject = "Registration Confirmation – INTERLINK 2026";
    htmlBody = buildCSPCEmailHtml(firstName, !!bgBlob);
  } else if (school === "UNAIR") {
    subject = "Registration Confirmation – INTERLINK 2026";
    htmlBody = buildUNAIREmailHtml(firstName, !!bgBlob);
  } else {
    // Unknown school — skip
    return;
  }

  mailOptions.htmlBody = htmlBody;

  GmailApp.sendEmail(recipientEmail, subject, "", mailOptions);
}


/**
 * Builds the HTML email body for CSPC students.
 *
 * @param {string} firstName - The registrant's first name.
 * @returns {string} HTML string.
 */
function buildCSPCEmailHtml(firstName, hasVirtualBg) {
  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>Registration Confirmation – INTERLINK 2026</title>
  <!-- Google Fonts: Inter -->
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
</head>
<body style="margin:0;padding:0;background-color:#F8F9FA;font-family:'Inter', 'Segoe UI', Arial, sans-serif; -webkit-font-smoothing: antialiased;">

  <table width="100%" cellpadding="0" cellspacing="0" style="background-color:#F8F9FA;padding:32px 0;">
    <tr>
      <td align="center">
        <table width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;background:#ffffff;border-radius:12px;overflow:hidden;box-shadow:0 1px 3px rgba(0, 0, 0, 0.04), 0 6px 24px rgba(0, 0, 0, 0.04);border: 1px solid #E2E8F0;">

          <!-- Header Banner Image -->
          <tr>
            <td style="padding:0;text-align:center;">
              <img src="https://lh3.googleusercontent.com/d/1JxkAPWuEBYhglR2jlLFDDLQmq7z0rKow" alt="Interlink 2026 Header" style="width:100%;max-width:600px;height:auto;display:block;border-bottom:1px solid #E2E8F0;" />
            </td>
          </tr>

          <!-- Body -->
          <tr>
            <td style="padding:40px 40px 0;">

              <p style="margin:0 0 8px;font-size:16px;color:#0F172A;font-weight:700;">
                Good day, ${firstName}!
              </p>

              <p style="margin:16px 0;font-size:14px;color:#334155;line-height:1.6;">
                Thank you for your interest in participating in this meaningful discussion!
                The <strong>College of Computer Studies</strong> proudly presents
                <strong>INTERLINK 2026: A CSPC-UNAIR Cross Campus Webinar</strong>.
                We are truly grateful for your interest and participation in this meaningful academic
                collaboration between <strong>Camarines Sur Polytechnic Colleges (CSPC)</strong> and
                <strong>Universitas Airlangga (UNAIR)</strong>.
              </p>

              <p style="margin:16px 0;font-size:14px;color:#334155;line-height:1.6;">
                This email serves as a reminder that <strong>INTERLINK 2026</strong> is fast approaching,
                and we are excited to have you with us for this event!
              </p>

            </td>
          </tr>

          <!-- Event Info Card -->
          <tr>
            <td style="padding:24px 40px;">
              <table width="100%" cellpadding="0" cellspacing="0" style="background:#F8FAFC;border-radius:8px;border-left:4px solid #0F172A;border-top:1px solid #E2E8F0;border-right:1px solid #E2E8F0;border-bottom:1px solid #E2E8F0;overflow:hidden;">
                <tr>
                  <td style="padding:24px;">

                    <p style="margin:0 0 16px;font-size:11px;font-weight:700;color:#0F172A;text-transform:uppercase;letter-spacing:1px;">
                      Event Details
                    </p>

                    <table cellpadding="0" cellspacing="0" width="100%">
                      <tr>
                        <td style="padding:6px 12px 6px 0;font-size:13px;color:#64748B;font-weight:600;white-space:nowrap;width:100px;">Date</td>
                        <td style="padding:6px 0;font-size:13px;color:#0F172A;font-weight:700;">${EVENT_DATE}</td>
                      </tr>
                      <tr>
                        <td style="padding:6px 12px 6px 0;font-size:13px;color:#64748B;font-weight:600;white-space:nowrap;">Platform</td>
                        <td style="padding:6px 0;font-size:13px;color:#1E293B;">Microsoft Teams </td>
                      </tr>
                      <tr>
                        <td style="padding:6px 12px 6px 0;font-size:13px;color:#64748B;font-weight:600;white-space:nowrap;">Time</td>
                        <td style="padding:6px 0;font-size:13px;color:#0F172A;font-weight:700;">9:30 AM – 4:00 PM <span style="color:#64748B;font-weight:400;">(GMT+8, Philippines)</span></td>
                      </tr>
                      <tr>
                        <td style="padding:6px 12px 6px 0;font-size:13px;color:#64748B;font-weight:600;white-space:nowrap;">Meeting ID</td>
                        <td style="padding:6px 0;font-size:13px;color:#1E293B;font-family:monospace;">${MEETING_ID}</td>
                      </tr>
                      <tr>
                        <td style="padding:6px 12px 6px 0;font-size:13px;color:#64748B;font-weight:600;white-space:nowrap;">Passcode</td>
                        <td style="padding:6px 0;font-size:13px;color:#0F172A;font-family:monospace;font-weight:700;">${MEETING_PASSCODE}</td>
                      </tr>
                    </table>

                  </td>
                </tr>
              </table>
            </td>
          </tr>

          <!-- Join Button -->
          <tr>
            <td style="padding:8px 40px 24px;text-align:center;">
              <a href="${MsTeams_LINK}"
                 style="display:inline-block;background-color:#0F172A;color:#ffffff;text-decoration:none;font-size:14px;font-weight:600;padding:12px 32px;border-radius:6px;letter-spacing:0.2px;box-shadow:0 1px 3px rgba(0,0,0,0.08);">
                Join MS Teams Meeting
              </a>
              <p style="margin:10px 0 0;font-size:12px;color:#64748B;">
                Or copy the link: <a href="${MsTeams_LINK}" style="color:#0F172A;text-decoration:underline;">${MsTeams_LINK}</a>
              </p>
            </td>
          </tr>

          <!-- Divider -->
          <tr>
            <td style="padding:0 40px;">
              <hr style="border:none;border-top:1px solid #E2E8F0;margin:0;"/>
            </td>
          </tr>

          <!-- Social Media -->
          <tr>
            <td style="padding:24px 40px;">
              <a href="${CSPC_FB_URL}"
                 style="display:inline-block;background-color:#1E293B;color:#ffffff;text-decoration:none;font-size:13px;font-weight:600;padding:10px 20px;border-radius:6px;border:1px solid #334155;">
                CSPC – College of Computer Studies Facebook
              </a>
            </td>
          </tr>

          <!-- Attachment Note -->
          <tr>
            <td style="padding:0 40px 24px;">
              <table width="100%" cellpadding="0" cellspacing="0" style="background:#F8FAFC;border-radius:6px;border-left:3px solid #64748B;border-top:1px solid #E2E8F0;border-right:1px solid #E2E8F0;border-bottom:1px solid #E2E8F0;">
                <tr>
                  <td style="padding:16px;font-size:13px;color:#475569;line-height:1.5;">
                    <strong>Note:</strong> Attached is the MS Teams virtual background for the event.
                    ${hasVirtualBg ? `
                    <div style="margin-top:12px;text-align:center;">
                      <img src="cid:virtual_bg" alt="MS Teams Virtual Background" style="width:100%;max-width:520px;height:auto;display:block;border-radius:4px;border:1px solid #E2E8F0;margin:0 auto;" />
                    </div>
                    ` : ""}
                  </td>
                </tr>
              </table>
            </td>
          </tr>

          <!-- Divider -->
          <tr>
            <td style="padding:0 40px;">
              <hr style="border:none;border-top:1px solid #E2E8F0;margin:0;"/>
            </td>
          </tr>

          <!-- Footer -->
          <tr>
            <td style="padding:28px 40px 36px;text-align:center;">
              <p style="margin:0 0 4px;font-size:13px;color:#64748B;line-height:1.5;">
                All the best,
              </p>
              <p style="margin:0;font-size:13px;font-weight:700;color:#0F172A;">
                ORGANIZERS
              </p>
              <p style="margin:2px 0 0;font-size:12px;color:#64748B;">
                International Faculty and Student Mobility Program 2025 Participants
              </p>
              <p style="margin:20px 0 0;font-size:11px;color:#94A3B8;">
                This is an automated email. Please do not reply directly to this message.
              </p>
            </td>
          </tr>

        </table>
      </td>
    </tr>
  </table>

</body>
</html>`;
}


/**
 * Builds the HTML email body for UNAIR students.
 *
 * @param {string} firstName - The registrant's first name.
 * @returns {string} HTML string.
 */
function buildUNAIREmailHtml(firstName, hasVirtualBg) {
  return `<!DOCTYPE html>
<html lang="id">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>Registration Confirmation – INTERLINK 2026</title>
  <!-- Google Fonts: Inter -->
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
</head>
<body style="margin:0;padding:0;background-color:#F8F9FA;font-family:'Inter', 'Segoe UI', Arial, sans-serif; -webkit-font-smoothing: antialiased;">

  <table width="100%" cellpadding="0" cellspacing="0" style="background-color:#F8F9FA;padding:32px 0;">
    <tr>
      <td align="center">
        <table width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;background:#ffffff;border-radius:12px;overflow:hidden;box-shadow:0 1px 3px rgba(0, 0, 0, 0.04), 0 6px 24px rgba(0, 0, 0, 0.04);border: 1px solid #E2E8F0;">

          <!-- Header Banner Image -->
          <tr>
            <td style="padding:0;text-align:center;">
              <img src="https://lh3.googleusercontent.com/d/1JxkAPWuEBYhglR2jlLFDDLQmq7z0rKow" alt="Interlink 2026 Header" style="width:100%;max-width:600px;height:auto;display:block;border-bottom:1px solid #E2E8F0;" />
            </td>
          </tr>

          <!-- Body -->
          <tr>
            <td style="padding:40px 40px 0;">

              <p style="margin:0 0 8px;font-size:16px;color:#0F172A;font-weight:700;">
                Selamat siang, ${firstName}!
              </p>

              <p style="margin:16px 0;font-size:14px;color:#334155;line-height:1.6;">
                Thank you for your interest in participating in this meaningful discussion!
                The <strong>College of Computer Studies</strong> proudly presents
                <strong>INTERLINK 2026: A CSPC-UNAIR Cross Campus Webinar</strong>.
                We are truly grateful for your interest and participation in this meaningful academic
                collaboration between <strong>Camarines Sur Polytechnic Colleges (CSPC)</strong> and
                <strong>Universitas Airlangga (UNAIR)</strong>.
              </p>

              <p style="margin:16px 0;font-size:14px;color:#334155;line-height:1.6;">
                This email serves as a reminder that <strong>INTERLINK 2026</strong> is fast approaching,
                and we are excited to have you with us for this event!
              </p>

            </td>
          </tr>

          <!-- Event Info Card -->
          <tr>
            <td style="padding:24px 40px;">
              <table width="100%" cellpadding="0" cellspacing="0" style="background:#F8FAFC;border-radius:8px;border-left:4px solid #0F172A;border-top:1px solid #E2E8F0;border-right:1px solid #E2E8F0;border-bottom:1px solid #E2E8F0;overflow:hidden;">
                <tr>
                  <td style="padding:24px;">

                    <p style="margin:0 0 16px;font-size:11px;font-weight:700;color:#0F172A;text-transform:uppercase;letter-spacing:1px;">
                      Event Details
                    </p>

                    <table cellpadding="0" cellspacing="0" width="100%">
                      <tr>
                        <td style="padding:6px 12px 6px 0;font-size:13px;color:#64748B;font-weight:600;white-space:nowrap;width:100px;">Date</td>
                        <td style="padding:6px 0;font-size:13px;color:#0F172A;font-weight:700;">${EVENT_DATE}</td>
                      </tr>
                      <tr>
                        <td style="padding:6px 12px 6px 0;font-size:13px;color:#64748B;font-weight:600;white-space:nowrap;">Platform</td>
                        <td style="padding:6px 0;font-size:13px;color:#1E293B;">Microsoft Teams</td>
                      </tr>
                      <tr>
                        <td style="padding:6px 12px 6px 0;font-size:13px;color:#64748B;font-weight:600;white-space:nowrap;">Time</td>
                        <td style="padding:6px 0;font-size:13px;color:#0F172A;font-weight:700;">8:30 AM – 3:00 PM <span style="color:#64748B;font-weight:400;">(WIB / GMT+7, Indonesia)</span></td>
                      </tr>
                      <tr>
                        <td style="padding:6px 12px 6px 0;font-size:13px;color:#64748B;font-weight:600;white-space:nowrap;">Meeting ID</td>
                        <td style="padding:6px 0;font-size:13px;color:#1E293B;font-family:monospace;">${MEETING_ID}</td>
                      </tr>
                      <tr>
                        <td style="padding:6px 12px 6px 0;font-size:13px;color:#64748B;font-weight:600;white-space:nowrap;">Passcode</td>
                        <td style="padding:6px 0;font-size:13px;color:#0F172A;font-family:monospace;font-weight:700;">${MEETING_PASSCODE}</td>
                      </tr>
                    </table>

                  </td>
                </tr>
              </table>
            </td>
          </tr>

          <!-- Join Button -->
          <tr>
            <td style="padding:8px 40px 24px;text-align:center;">
              <a href="${MsTeams_LINK}"
                 style="display:inline-block;background-color:#0F172A;color:#ffffff;text-decoration:none;font-size:14px;font-weight:600;padding:12px 32px;border-radius:6px;letter-spacing:0.2px;box-shadow:0 1px 3px rgba(0,0,0,0.08);">
                Join MS Teams Meeting
              </a>
              <p style="margin:10px 0 0;font-size:12px;color:#64748B;">
                Or copy the link: <a href="${MsTeams_LINK}" style="color:#0F172A;text-decoration:underline;">${MsTeams_LINK}</a>
              </p>
            </td>
          </tr>

          <!-- Divider -->
          <tr>
            <td style="padding:0 40px;">
              <hr style="border:none;border-top:1px solid #E2E8F0;margin:0;"/>
            </td>
          </tr>

          <!-- Social Media -->
          <tr>
            <td style="padding:24px 40px;">
              <a href="${CSPC_FB_URL}"
                 style="display:inline-block;background-color:#1E293B;color:#ffffff;text-decoration:none;font-size:13px;font-weight:600;padding:10px 20px;border-radius:6px;border:1px solid #334155;">
                CSPC – College of Computer Studies Facebook
              </a>
            </td>
          </tr>

          <!-- Attachment Note -->
          <tr>
            <td style="padding:0 40px 24px;">
              <table width="100%" cellpadding="0" cellspacing="0" style="background:#F8FAFC;border-radius:6px;border-left:3px solid #64748B;border-top:1px solid #E2E8F0;border-right:1px solid #E2E8F0;border-bottom:1px solid #E2E8F0;">
                <tr>
                  <td style="padding:16px;font-size:13px;color:#475569;line-height:1.5;">
                    <strong>Note:</strong> Attached is the MS Teams virtual background for the event.
                    ${hasVirtualBg ? `
                    <div style="margin-top:12px;text-align:center;">
                      <img src="cid:virtual_bg" alt="MS Teams Virtual Background" style="width:100%;max-width:520px;height:auto;display:block;border-radius:4px;border:1px solid #E2E8F0;margin:0 auto;" />
                    </div>
                    ` : ""}
                  </td>
                </tr>
              </table>
            </td>
          </tr>

          <!-- Divider -->
          <tr>
            <td style="padding:0 40px;">
              <hr style="border:none;border-top:1px solid #E2E8F0;margin:0;"/>
            </td>
          </tr>

          <!-- Footer -->
          <tr>
            <td style="padding:28px 40px 36px;text-align:center;">
              <p style="margin:0 0 4px;font-size:13px;color:#64748B;line-height:1.5;">
                All the best,
              </p>
              <p style="margin:0;font-size:13px;font-weight:700;color:#0F172A;">
                ORGANIZERS
              </p>
              <p style="margin:2px 0 0;font-size:12px;color:#64748B;">
                International Faculty and Student Mobility Program 2025 Participants
              </p>
              <p style="margin:20px 0 0;font-size:11px;color:#94A3B8;">
                This is an automated email. Please do not reply directly to this message.
              </p>
            </td>
          </tr>

        </table>
      </td>
    </tr>
  </table>

</body>
</html>`;
}


/**
 * TEST FUNCTION — Run this BEFORE deploying to verify
 * that emails look correct in your inbox.
 *
 * Steps:
 *   1. Open this file in the Google Apps Script Editor.
 *   2. In the top toolbar, select "testSendEmail" from the function dropdown.
 *   3. Click Run.
 *   4. Check the inbox of the email configured in RECIPIENT below.
 */
function testSendEmail() {
  // Change these two values to your own test details
  const TEST_RECIPIENT = "your-test-email@example.com"; // Where to send the test
  const TEST_FIRST_NAME = "Juan";                         // Name to use in greeting

  // Test CSPC email
  const cspcPayload = {
    email: TEST_RECIPIENT,
    firstName: TEST_FIRST_NAME,
    school: "CSPC"
  };

  // Test UNAIR email
  const unairPayload = {
    email: TEST_RECIPIENT,
    firstName: TEST_FIRST_NAME,
    school: "UNAIR"
  };

  try {
    sendInvitationEmail(cspcPayload);
    Logger.log("[OK] CSPC test email sent to: " + TEST_RECIPIENT);
  } catch (err) {
    Logger.log("[FAIL] CSPC email failed: " + err.toString());
  }

  try {
    sendInvitationEmail(unairPayload);
    Logger.log("[OK] UNAIR test email sent to: " + TEST_RECIPIENT);
  } catch (err) {
    Logger.log("[FAIL] UNAIR email failed: " + err.toString());
  }
}


/**
 * Helper function to generate JSON output.
 * Note: Google Apps Script Web Apps automatically handle CORS headers
 * (Access-Control-Allow-Origin: *) for requests.
 */
function createJsonResponse(success, message) {
  return ContentService
    .createTextOutput(
      JSON.stringify({
        success: success,
        message: message
      })
    )
    .setMimeType(ContentService.MimeType.JSON);
}
