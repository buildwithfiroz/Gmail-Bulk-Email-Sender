function sendEmails() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  var data = sheet.getDataRange().getValues();
  var redirectUrl = '';
  var scriptUrl = "";    // Update if needed

  var quota = MailApp.getRemainingDailyQuota();
  var sentCount = 0;

  // Batch buffers so we don't call setValue() per-cell per-row
  var trackingIdCol = [];   // Column G
  var sentStatusCol = [];   // Column L (new "Sent" flag, adjust index as needed)

  for (var i = 1; i < data.length; i++) {
    var rowNum = i + 1;
    var alreadySent = data[i][11]; // Column L - adjust to your actual "Sent" column
    trackingIdCol.push([data[i][6]]);   // preserve existing tracking ID by default
    sentStatusCol.push([alreadySent || ""]);

    if (alreadySent === "Yes") continue; // skip rows already sent

    try {
      var name = data[i][0];
      var email = data[i][1];
      var cc1 = data[i][2];
      var cc2 = data[i][3];
      var subject = data[i][4];
      var destinationUrl = data[i][5];
      var htmlContent = data[i][10];

      if (!email) continue;
      if (!htmlContent || typeof htmlContent !== 'string') {
        throw new Error("Invalid or empty HTML content in column K");
      }
      if (quota <= sentCount) {
        Logger.log("Daily mail quota reached. Stopping at row " + rowNum);
        break;
      }

      var trackingId = Utilities.getUuid();
      var trackingPixelUrl = scriptUrl + "?id=" + trackingId + "&action=open";
      var clickTrackingUrl = scriptUrl + "?id=" + trackingId + "&action=click&redirect=" + encodeURIComponent(destinationUrl);

      var body = htmlContent
        .replace(/{{click-tracking-url}}/g, clickTrackingUrl)
        .replace(/{{name}}/g, name)
        .replace(/{{tracking-pixel-url}}/g, trackingPixelUrl);

      var ccEmails = [cc1, cc2].filter(Boolean).join(",");

      MailApp.sendEmail({
        to: email,
        cc: ccEmails || "",
        subject: subject,
        htmlBody: body
      });

      trackingIdCol[i - 1] = [trackingId];
      sentStatusCol[i - 1] = ["Yes"];
      sentCount++;
      Logger.log("Email sent to: " + name);
    } catch (error) {
      Logger.log("Error sending email to row " + rowNum + ": " + error.toString());
    }
  }

  // Single batched write instead of per-row setValue calls
  sheet.getRange(2, 7, trackingIdCol.length, 1).setValues(trackingIdCol);   // Col G
  sheet.getRange(2, 12, sentStatusCol.length, 1).setValues(sentStatusCol); // Col L
}
