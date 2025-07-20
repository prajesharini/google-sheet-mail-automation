// 📧 Bulk Email Menu
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("📧 Bulk Mail")
    .addItem("Send Emails", "sendBulkEmails")
    .addToUi();

  ui.createMenu("📎 Attachment")
    .addItem("Upload File to Cell", "showUploadDialog")
    .addToUi();
}

// 📧 Bulk Email Function
function sendBulkEmails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    SpreadsheetApp.getUi().alert("⚠️ No data found to send emails.");
    return;
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues(); // A to E
  let count = 0;

  data.forEach((row, index) => {
    const [To, From, Subject, Body, AttachmentInput] = row;
    if (To && Subject && Body) {
      let attachments = [];

      if (AttachmentInput) {
        try {
          const fileId = extractFileId(AttachmentInput);
          const file = DriveApp.getFileById(fileId);
          attachments.push(file.getBlob());
        } catch (err) {
          Logger.log(`❌ Error at row ${index + 2} (Attachment): ${err}`);
        }
      }

      const fullBody = `From: ${From || "N/A"}\n\n${Body}`;

      try {
        GmailApp.sendEmail(To, Subject, fullBody, {
          ...(attachments.length > 0 ? { attachments } : {})
        });
        count++;
      } catch (e) {
        Logger.log(`❌ Failed to send email at row ${index + 2}: ${e}`);
      }
    }
  });

  SpreadsheetApp.getUi().alert("✅ " + count + " email(s) sent successfully!");
}

// 📎 Open HTML Upload Dialog
function showUploadDialog() {
  const html = HtmlService.createHtmlOutputFromFile("FileUploader")
    .setWidth(400)
    .setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(html, "📎 Upload the document");
}

// 📎 Upload file to Drive and insert URL
function uploadToDrive(filename, base64Data) {
  const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), undefined, filename);
  const file = DriveApp.getRootFolder().createFile(blob);
  const fileUrl = file.getUrl();

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getActiveCell().setValue(fileUrl);

  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Uploads Log");
  if (logSheet) {
    logSheet.appendRow([new Date(), filename, fileUrl]);
  }

  return fileUrl;
}

// 📎 Extract File ID from Google Drive URL
function extractFileId(input) {
  const idPattern = /[-\w]{25,}/;
  const match = input.match(idPattern);
  if (match && match[0]) {
    return match[0];
  } else {
    throw new Error("Invalid Google Drive file link or ID");
  }
}
function doGet(e) {
  return HtmlService.createHtmlOutput("Web App is working!");
}