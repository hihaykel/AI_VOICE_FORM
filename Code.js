/**
 * GOOGLE SHEETS AUTOMATION WITH GEMINI AI
 * 
 * Setup Instructions:
 * 1. Replace "YOUR_GEMINI_API_KEY" with your own key from Google AI Studio.
 * 2. Configure an Installed Trigger in Apps Script:
 *    - Click on the clock icon (Triggers) on the left menu.
 *    - Click "Add Trigger".
 *    - Choose which function to run: "automatisationGeminiForm".
 *    - Select event type: "On edit".
 */

function automatisationGeminiForm(e) {
  if (!e) return;
  
  var range = e.range;
  var sheet = range.getSheet();
  
  // Strictly runs on the input cell F4 of 'Sheet1'
  if (sheet.getName() === "Sheet1" && range.getA1Notation() === "F4") {
    var input = range.getValue().toString().trim();
    if (input === "") return;
    
    // Set your Gemini API Key here
    var apiKey = "Gemini_API_Key"; 
    var url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=" + apiKey;
    
    // Prompt engineered to request strict JSON extraction
    var prompt = "Extract the following details from this text: Title (like Mr, Mrs, Ms, Dr), Full Name, Email, and Phone Number. " +
                 "Text: '" + input + "'. " +
                 "Respond ONLY with a valid JSON object containing these keys: 'title', 'fullName', 'email', 'phone'. " +
                 "Do not include any markdown formatting like ```json or wrappers. Just the raw JSON.";
    
    var payload = {
      "contents": [{ "parts": [{ "text": prompt }] }],
      "generationConfig": { 
        "temperature": 0.1,
        "responseMimeType": "application/json" // Forces the API to output native JSON structure
      }
    };
    
    var options = {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };
    
    try {
      var response = UrlFetchApp.fetch(url, options);
      
      if (response.getResponseCode() === 200) {
        var jsonText = response.getContentText().trim();
        
        // Safety cleanup if the model accidentally wraps text in Markdown blocks
        jsonText = jsonText.replace(/^```json\s*/i, "").replace(/```$/, "").trim();
        
        var aiResult = JSON.parse(jsonText);
        
        // Anti-parsing error fix (#ERROR!) for phone numbers starting with "+"
        var phoneVal = aiResult.phone ? aiResult.phone.toString() : "";
        if (phoneVal.startsWith("+")) {
          phoneVal = "'" + phoneVal;
        }
        
        // Populating the UI form fields
        sheet.getRange("C2").setValue(aiResult.title || "");
        sheet.getRange("C4").setValue(aiResult.fullName || "");
        sheet.getRange("C6").setValue(aiResult.email || "");
        sheet.getRange("C8").setValue(phoneVal);
        
      } else {
        SpreadsheetApp.getUi().alert("Gemini API Error: " + response.getContentText());
      }
    } catch (err) {
      SpreadsheetApp.getUi().alert("Execution Error: " + err.toString());
    }
  }
}

/**
 * BLUE BUTTON ACTION: Saves historical data to Sheet2 and resets the form UI
 */
function saveAndClearForm() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formSheet = ss.getSheetByName("Sheet1"); 
  var dataSheet = ss.getSheetByName("Sheet2"); 
  
  // Automatically creates Sheet2 if it doesn't exist yet
  if (!dataSheet) {
    dataSheet = ss.insertSheet("Sheet2");
  }
  
  var title = formSheet.getRange("C2").getValue().toString();
  var fullName = formSheet.getRange("C4").getValue().toString();
  var email = formSheet.getRange("C6").getValue().toString();
  var phoneValue = formSheet.getRange("C8").getValue().toString();
  
  // Minimum validation: stops the process if the form is empty
  if (title === "" && fullName === "") {
    SpreadsheetApp.getUi().alert("The form is completely empty!");
    return;
  }
  
  // Double safety measure for phone format right before logging into Sheet2 row
  if (phoneValue.startsWith("+")) {
    phoneValue = "'" + phoneValue;
  }

  // Appending data to the database tab
  var values = [new Date(), title, fullName, email, phoneValue];
  dataSheet.appendRow(values); 
  
  // Forces Google Sheets to refresh UI rendering instantly
  SpreadsheetApp.flush(); 
  
  // Complete form UI cleanup
  formSheet.getRange("C2").clearContent();
  formSheet.getRange("C4").clearContent();
  formSheet.getRange("C6").clearContent();
  formSheet.getRange("C8").clearContent();
  formSheet.getRange("F4").clearContent(); // Clears voice input field
  
  // Confirmation popup
  SpreadsheetApp.getUi().alert("Gemini AI successfully processed your voice and saved the data!");
}
