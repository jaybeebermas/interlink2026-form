/**
 * Deploys as a Google Apps Script Web App.
 * Intercepts POST requests from the frontend, validates fields, check for duplicates,
 * and appends registration data to the spreadsheet.
 */

function doPost(e) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName("Interlink2026");

    // Dynamic Sheet Initialization: If the target sheet tab doesn't exist, create it with headers
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
      // Format headers: Bold fonts, centered text, freeze first row
      const headerRange = sheet.getRange(1, 1, 1, 12);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#f1f5f9");
      headerRange.setFontColor("#0f172a");
      sheet.setFrozenRows(1);
    }

    // Parse JSON payload from POST request
    const data = JSON.parse(e.postData.contents);

    // Backend validation: Ensure required elements are present
    const email = (data.email || "").trim().toLowerCase();
    const firstName = (data.firstName || "").trim();
    const lastName = (data.lastName || "").trim();
    const school = (data.school || "").trim();
    const yearLevel = (data.yearLevel || "").trim();

    if (!email || !firstName || !lastName || !school || !yearLevel) {
      return createJsonResponse(false, "Required fields are missing. Please fill in the entire form.");
    }

    // Duplicate Check: Look up email to prevent double registrations
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) { // Skip header row
      const existingEmail = rows[i][4].toString().trim().toLowerCase(); // Index 4 = Column E (Email)
      if (existingEmail === email) {
        return createJsonResponse(false, `The email address "${data.email}" is already registered.`);
      }
    }

    // Write row to active sheet
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
      new Date() // Timestamp column
    ]);

    return createJsonResponse(true, "Registration successful! Welcome to Interlink 2026.");

  } catch(error) {
    return createJsonResponse(false, "Server Error: " + error.toString());
  }
}

/**
 * Helper function to generate JSON output.
 * Note: Google Apps Script Web Apps automatically handle and attach CORS headers 
 * (Access-Control-Allow-Origin: *) for requests, so we do not need to manually set them.
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
