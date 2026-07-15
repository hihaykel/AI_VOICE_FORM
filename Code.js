// ====== CONFIGURATION ======
const API_KEY = "your API_KEY";
const GEMINI_URL = "https://generativelanguage.googleapis.com/v1/models/gemini-2.5-flash:generateContent?key=" + API_KEY;

/**
 * Fonction automatique déclenchée lors de la modification de la feuille
 */
function onEdit(e) {
  // Sécurité essentielle pour éviter l'erreur de ta capture d'écran image_8c32c0.png
  if (!e || !e.range) return;
  
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  
  // On vérifie qu'on est sur l'onglet "Orders" et sur la cellule B2
  if (sheet.getName() !== "Orders" || range.getA1Notation() !== "B2") return;
  
  const voiceQuery = range.getValue().toString().trim();
  
  // Si on efface le texte de B2, on nettoie la réponse en F2
  if (voiceQuery === "") {
    sheet.getRange("F2").clearContent();
    return;
  }
  
  // Lancement du processus
  sheet.getRange("F2").setValue("🔍 Searching...");
  SpreadsheetApp.flush(); 
  
  const prompt = "Analyze this search query: '" + voiceQuery + "'. " +
                 "Extract filters for a sales database. Available columns to filter by: 'Customer Name', 'State', 'Category'. " +
                 "Also identify what numeric column the user wants to calculate: 'Sales' or 'Profit'. " +
                 "Return ONLY a clean JSON object with keys: " +
                 "'customer' (string or null), 'state' (string or null), 'category' (string or null), 'metric' ('Sales' or 'Profit'). " +
                 "Do not wrap the response in markdown code blocks like ```json ... ```. Output raw JSON string only.";
                 
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
      sheet.getRange("F2").setValue("⚠️ API Error");
      return;
    }
    
    let rawJsonText = jsonResponse.candidates[0].content.parts[0].text.trim();
    
    if (rawJsonText.startsWith("```")) {
      rawJsonText = rawJsonText.replace(/^```json/, "").replace(/^```/, "").replace(/```$/, "").trim();
    }
    
    const criteria = JSON.parse(rawJsonText);
    const result = searchAndCalculate(sheet, criteria);
    
    // Affichage final en F2
    sheet.getRange("F2").setValue(result);
    
  } catch (error) {
    sheet.getRange("F2").setValue("⚠️ Error");
  }
}

/**
 * Fonction interne de recherche (Ligne 6 = en-têtes, donc index 5)
 */
function searchAndCalculate(sheet, criteria) {
  const rows = sheet.getDataRange().getValues();
  
  // SEULE MODIFICATION : 5 au lieu de 3 pour cibler la ligne 6 (vos en-têtes)
  const idxHeaderRow = 5; 
  
  if (rows.length <= idxHeaderRow + 1) return "No data found.";
  
  const headers = rows[idxHeaderRow];
  const idxCustomer = headers.indexOf("Customer Name");
  const idxState = headers.indexOf("State");
  const idxCategory = headers.indexOf("Category");
  const idxMetric = headers.indexOf(criteria.metric || "Sales");
  
  if (idxMetric === -1) return "Column not found.";
  
  let total = 0;
  let matchCount = 0;
  
  for (let i = idxHeaderRow + 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row[0]) continue; 
    
    let match = true;
    
    if (criteria.customer && row[idxCustomer].toString().toLowerCase().indexOf(criteria.customer.toLowerCase()) === -1) {
      match = false;
    }
    if (criteria.state && row[idxState].toString().toLowerCase().indexOf(criteria.state.toLowerCase()) === -1) {
      match = false;
    }
    if (criteria.category && row[idxCategory].toString().toLowerCase().indexOf(criteria.category.toLowerCase()) === -1) {
      match = false;
    }
    
    if (match) {
      let rawValue = row[idxMetric].toString().replace(/\s/g, "").replace(",", ".");
      const value = parseFloat(rawValue);
      if (!isNaN(value)) {
        total += value;
      }
      matchCount++;
    }
  }
  
  if (matchCount === 0) return "No matching entries.";
  
  const formattedTotal = total.toLocaleString('en-US', { style: 'currency', currency: 'USD' });
  return "Total " + criteria.metric + ": " + formattedTotal + " (" + matchCount + " orders)";
}
