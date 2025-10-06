# Auto Sync Google Form Dropdown

> An automation script that syncs dropdown options from Google Sheets directly into Google Forms, enabling bulk imports and seamless updates.

---

## 📊 Overview

**Project Type:** Automation

**Problem:** When users create forms in Google Forms with dropdown questions, they typically add options manually, one by one. This becomes a significant pain point when dealing with high-volume option lists that require frequent updates.

**Solution:** An automation script that syncs dropdown options from Google Sheets directly into Google Forms, enabling bulk imports and real-time updates.

**Impact:** This script allows users to organize their dropdown options in a structured spreadsheet and eliminates manual effort when updating options across multiple forms.

---

## ✨ Key Features

- Form Sync Button
- Form Sync Status Sheet

---

## 🔧 Tech Stack

**Core:**
- App Script
- Google Sheet

---

## 🚀 Quick Start
1. Prepare your Google Sheet
    - Create a dedicated worksheet to organize your form option lists
    - Structure your data with clear column headers
2. Configure your Google Form
    - Ensure your target dropdown questions are named correctly (these names will be used for mapping)
    - Set question type to "Dropdown"
    - Leave the options blank (they will be populated automatically)
3. Access Google Apps Script
    - Open your Google Sheet
    - Navigate to Extensions > Apps Script
4. Deploy the automation script
    - Copy the provided code into the Apps Script editor ➡️ auto_sync_script.gs
    - Customize the following variables:
        - FORM_ID: Your Google Form ID (found in the form URL)
        - SHEET_ID: Your Google Sheet ID (found in the sheet URL)
        - DROPDOWN: Match your worksheet name and column name to the corresponding form questions
    - Save the script (Ctrl+S or Cmd+S)
5. Activate the sync menu
    - Refresh your Google Sheet
    - You'll see a new menu "🔄 Form Sync" in the menu bar
    - Click 🔄 Form Sync > Sync Dropdowns Now to run the sync anytime

---

## 📁 Project Structure

```
Auto Sync Google Form Dropdown/
├── data/              # Data files
├── src/               # Source code
├── tests/             # Test files
├── requirements.txt   # Dependencies
└── README.md         # This file
```

---

## 📄 License

MIT License - see [LICENSE](LICENSE)

---

## 👤 Author

**Your Name**
- GitHub: https://github.com/igarahmawati
- LinkedIn: https://linkedin.com/in/iga-rahmawati
- Email: hi.igarwt@gmail.com

---

## 🙏 Acknowledgments

- Data source: [G.Sheet](https://docs.google.com/spreadsheets/d/1s31PTo43iLpoo6dTgmDkvEqjCFWjd8cF2sIAq4sOrLs)
