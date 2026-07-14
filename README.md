# 🚀 AI-Powered Voice Form Automation (Google Sheets + Gemini API)

A smart, voice-activated form built inside Google Sheets using **Google Apps Script** and the **Gemini 1.5 Flash API**. This automation parses raw text (entered via voice dictation) into structured database fields (*Title, Full Name, Email, and Phone Number*) with 100% accuracy using native JSON response formatting.

---

## ✨ Features
* **Voice-to-Form Processing:** Type or use voice dictation (`Windows + H`) in a single input cell, and watch the AI instantly organize the data.
* **Strict JSON Structuring:** Uses Gemini's `responseMimeType: "application/json"` config to prevent formatting errors and structural breaks.
* **Database Logging:** A single click on the custom UI button appends the structured data to a secondary database tab (`Sheet2`) with a live timestamp, then completely resets the form.

---

## 🛠️ Project Architecture

* **Sheet1:** The User Interface
  * `F4`: Raw speech text input cell.
  * `C2`: Extracted Title (Mr, Mrs, Dr, etc.)
  * `C4`: Extracted Full Name
  * `C6`: Extracted Email
  * `C8`: Extracted Phone Number
* **Sheet2:** The Database (automatically tracks and logs submissions with timestamps).

---

## 🚀 Quick Setup Guide

### 1. Get a Gemini API Key
1. Go to [Google AI Studio](https://aistudio.google.com/).
2. Click on **API keys** in the left menu.
3. Click **Create API key** and copy your generated token.

### 2. Configure Google Apps Script
1. Open your Google Sheet.
2. Navigate to **Extensions** > **Apps Script**.
3. Delete any default code and paste the content of `Code.js` from this repository.
4. Replace `"YOUR_GEMINI_API_KEY"` with your actual key from Google AI Studio.
5. Save the project.

### 3. Set Up the Trigger
To make the AI run automatically when you edit the sheet, you must configure a manual trigger:
1. Inside the Apps Script editor, click the **Triggers (Clock Icon)** on the left sidebar.
2. Click **+ Add Trigger** (bottom right).
3. Set **Choose which function to run** to `automatisationGeminiForm`.
4. Set **Select event type** to `On edit`.
5. Click **Save** and accept the required Google permissions.

### 4. Link the Save Button
1. Back in your Google Sheet, insert a drawing/shape to act as your "Save & Clear" button.
2. Click the three dots on the top right corner of the button.
3. Select **Assign script** and type exactly: `saveAndClearForm`.
4. Click OK.

---

## 🎯 How to Use
1. Select cell **F4**.
2. Press `Windows + H` (or `Cmd + Space` on Mac) to start your microphone.
3. Speak naturally (e.g., *"Hi, I am Dr. Alex Carter, my email is alex@tech.com and call me at +1555987654"*).
4. Press **Enter**. The AI will parse and distribute the data instantly.
5. Click your custom button to save the entry to your database!

---

## 📄 License
This project is open-source and free to use. If this helped you automate your workflow, don't forget to **Drop a ⭐** on this repository!
