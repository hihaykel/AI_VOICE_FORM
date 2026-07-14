// ====== CONFIGURATION ======
const API_KEY = "AQ.....API_KEY";
// Utilisation du modèle stable mis à jour gemini-2.5-flash pour corriger l'erreur 404
const GEMINI_URL = "https://generativelanguage.googleapis.com/v1/models/gemini-2.5-flash:generateContent?key=" + API_KEY;

/**
 * Fonction déclenchée par le Trigger "On edit" lors de la modification de F4
 */
function automatisationGeminiForm(e) {
  if (!e) return;
  
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  
  if (sheet.getName() !== "Sheet1" || range.getA1Notation() !== "F4") return;
  
  const voiceInput = range.getValue().toString().trim();
  if (voiceInput === "") return;
  
  // On demande explicitement une structure JSON pure dans le prompt
  const prompt = "Extract the following entities from this text: Title (Mr, Mrs, Dr, etc.), Full Name, Email Address, and Phone Number. " +
                 "Return ONLY a clean JSON object with keys: 'title', 'name', 'email', 'phone'. " +
                 "Do not wrap the response in markdown code blocks like ```json ... ```. Output raw JSON string only. " +
                 "Text to analyze: " + voiceInput;
                 
  const payload = {
    "contents": [{
      "parts": [{
        "text": prompt
      }]
    }]
  };
  
  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };
  
  try {
    const response = UrlFetchApp.fetch(GEMINI_URL, options);
    const jsonResponse = JSON.parse(response.getContentText());
    
    if (jsonResponse.error) {
      SpreadsheetApp.getUi().alert("Gemini API Error: " + JSON.stringify(jsonResponse.error, null, 2));
      return;
    }
    
    // Extraction du texte généré
    let rawJsonText = jsonResponse.candidates[0].content.parts[0].text.trim();
    
    // Nettoyage de sécurité si l'IA ajoute quand même des balises Markdown
    if (rawJsonText.startsWith("```")) {
      rawJsonText = rawJsonText.replace(/^```json/, "").replace(/^```/, "").replace(/```$/, "").trim();
    }
    
    const data = JSON.parse(rawJsonText);
    
    // Remplissage des cases du formulaire
    sheet.getRange("C2").setValue(data.title || "");
    sheet.getRange("C4").setValue(data.name || "");
    sheet.getRange("C6").setValue(data.email || "");
    sheet.getRange("C8").setValue(data.phone || "");
    
  } catch (error) {
    SpreadsheetApp.getUi().alert("Script Error: " + error.toString());
  }
}

/**
 * Fonction assignée au bouton bleu (Sauvegarde et Réinitialisation)
 */
function saveAndClearForm() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet1 = ss.getSheetByName("Sheet1");
  const sheet2 = ss.getSheetByName("Sheet2");
  
  const title = sheet1.getRange("C2").getValue();
  const name = sheet1.getRange("C4").getValue();
  const email = sheet1.getRange("C6").getValue();
  const phone = sheet1.getRange("C8").getValue();
  
  if (!title && !name && !email && !phone) {
    SpreadsheetApp.getUi().alert("The form is empty!");
    return;
  }
  
  const timestamp = new Date();
  sheet2.appendRow([timestamp, title, name, email, phone]);
  
  sheet1.getRange("F4").clearContent();
  sheet1.getRange("C2").clearContent();
  sheet1.getRange("C4").clearContent();
  sheet1.getRange("C6").clearContent();
  sheet1.getRange("C8").clearContent();
  
  SpreadsheetApp.getUi().alert("Data successfully saved to Sheet2!");
}
