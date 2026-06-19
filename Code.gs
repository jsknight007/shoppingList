/**
 * BACKEND LOGIC: Handles Add, Update, and Delete actions from the UI
 */
function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = JSON.parse(e.postData.contents);
  
  // ADD ITEM
  if (data.action === "add") {
    let rawText = data.text.trim();
    let parts = rawText.split(" ");
    let name = parts[0]; 
    let item = parts.slice(1).join(" "); 
    if (!item) { item = name; name = "General"; }
    sheet.appendRow([new Date(), name, item, ""]); 
    return ContentService.createTextOutput("Success");
  }
  
  // UPDATE STORE
  if (data.action === "update") {
    sheet.getRange(data.index, 4).setValue(data.value);
    return ContentService.createTextOutput("Updated");
  }

  // DELETE
  if (data.action === "delete") {
    const rows = sheet.getDataRange().getValues();
    for (let i = rows.length - 1; i >= 1; i--) {
      if (rows[i][1] === data.name && rows[i][2] === data.item) {
        sheet.deleteRow(i + 1); 
        break; 
      }
    }
    return ContentService.createTextOutput("Deleted");
  }

  // ON-DEMAND EMAIL
  if (data.action === "email_demand") {
    const wifeEmail = "jsknight007@gmail.com";
    const body = `Items for ${data.store}:\n\n${data.items}`;
    MailApp.sendEmail(wifeEmail, `🛒 Shopping List: ${data.store}`, body);
    return ContentService.createTextOutput("Email Sent");
  }
}

/**
 * FETCH LOGIC: Supplies the UI with the current list
 */
function doGet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const fullData = sheet.getDataRange().getValues();
  
  if (fullData.length <= 1) {
    return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
  }

  const rows = fullData.slice(1);
  const list = rows.map((row, index) => {
    let dateVal = row[0];
    let formattedDate = "N/A";

    // Date Parsing Logic
    try {
      if (dateVal instanceof Date) {
        formattedDate = Utilities.formatDate(dateVal, "GMT-6", "MM/dd");
      } else if (dateVal && dateVal !== "") {
        // If it's a string, try to convert it to a Date first
        let convertedDate = new Date(dateVal);
        if (!isNaN(convertedDate.getTime())) {
          formattedDate = Utilities.formatDate(convertedDate, "GMT-6", "MM/dd");
        }
      }
    } catch (e) {
      formattedDate = "Err";
    }

    return {
      id: index + 2, 
      date: formattedDate,
      name: row[1],
      item: row[2],
      store: row[3] || "" 
    };
  });

  return ContentService.createTextOutput(JSON.stringify(list)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * AUTOMATION: Groups by store and emails your wife
 * Set this up in Triggers (Clock icon)
 */
function sendDailySummary() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return; 

  const rows = data.slice(1);
  const stores = {};

  rows.forEach(row => {
    const store = row[3] || "Unassigned/New";
    if (!stores[store]) stores[store] = [];
    stores[store].push(`${row[1]}: ${row[2]}`);
  });

  let emailBody = "<div style='font-family:sans-serif;'><h2>🛒 Today's Needs List</h2>";
  for (const store in stores) {
    emailBody += `<h3 style='color:#2b7de9;'>📍 ${store}</h3><ul>`;
    stores[store].forEach(item => {
      emailBody += `<li style='margin-bottom:5px;'>${item}</li>`;
    });
    emailBody += "</ul>";
  }
  emailBody += "</div>";

  MailApp.sendEmail({
    to: "jsknight007@gmail.com", // <-- UPDATE THIS EMAIL
    subject: "Grocery & Needs Summary",
    htmlBody: emailBody
  });
}
