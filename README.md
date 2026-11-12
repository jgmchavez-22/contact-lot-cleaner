# Contact Lot Cleaner // Telephony & CRM Lead Normalizer

> An automated dual-interface utility (Python CLI + Client-Side Web Application) that standardizes B2B contact lists, normalizes phone numbers for auto-dialers, infers geographic regions, splits full names, and screens job titles against custom Do-Not-Call (DNC) exclusion lists.

![Python](https://img.shields.io/badge/Python-3.9+-3776AB?style=flat&logo=python&logoColor=white)
![JavaScript](https://img.shields.io/badge/JavaScript-ES6+-F7DF1E?style=flat&logo=javascript&logoColor=black)
![SheetJS](https://img.shields.io/badge/SheetJS-0.18.5-green?style=flat)
![License](https://img.shields.io/badge/License-MIT-blue.svg)

---

## 📋 Overview

### The Problem
Operational sales and outbound calling campaigns require clean, standardized lead lists formatted specifically for CRM dialer engines. Raw lead exports frequently suffer from:
1. **Unformatted Phone Numbers:** Inconsistent punctuation, missing country codes, or local prefixes cause auto-dialers to fail or trigger costly misdials.
2. **Unsplit Names:** Combined "Full Name" strings break CRM personalization tokens and email outreach templates.
3. **Non-Target Job Titles:** Leads containing excluded roles (e.g., interns, assistants, students) waste outbound agent hours.
4. **Mismatched Column Headers:** Manual column realignment and region tagging take hours per campaign.

### The Solution
**Contact Lot Cleaner** automates this end-to-end normalization pipeline in an 8-stage sequence, offered both as a high-speed Python terminal utility and as a zero-dependency in-browser HTML tool powered by SheetJS.

---

## ⚡ 8-Stage Normalization Pipeline

1. **Dialer Prefix Ingestion:** Automatically injects `Upload` (`Yes`) and `Result Code` (`New`) columns for all non-empty contact rows.
2. **Name Decomposition:** Parses `Full Name` into dedicated `First Name` and `Last Name` fields.
3. **Country-Specific Phone Validation:** Validates contact numbers against target country codes using `libphonenumber` / `phonenumbers`. Discards out-of-region numbers unless it is the contact's sole phone record (flagged as `Number outside specified country`).
4. **Auto-Dialer Normalization:** Strips formatting symbols (`+`, `(`, `)`, `-`, spaces) to yield pure digit sequences with prefixed country codes across up to 3 priority lines (`Phone`, `Phone 2`, `Phone3`).
5. **Geographic Region Mapping:** Automatically maps country names to operational regional territories (e.g., `SEA`, `ANZ`, `Europe`, `North America`, `MEA`).
6. **Job Title Screening:** Matches contact titles against custom exclusion lists (`excluded_titles.txt`), automatically tagging disqualified leads as `DNC - Not Target Job Title`.
7. **Multi-Key Sorting:** Sorts records deterministically across `Phone`, `Phone 2`, `Phone3`, `Company Name`, and `Country`.
8. **Dialer Schema Standardisation:** Re-orders columns into production-ready CRM sequence.

---

## 🏛️ Repository Structure

```
contact-lot-cleaner/
├── contact_lot_cleaner.py     # Standalone Python CLI pipeline (Pandas & Phonenumbers)
├── contact_lot_cleaner.html   # Zero-install client-side web application (SheetJS)
├── excluded_titles.txt        # Custom DNC / disqualified job title filter list
├── libphonenumber-max.js      # Full international telephony library for web runtime
├── xlsx.full.min.js           # Client-side workbook parser & generator
├── requirements.txt           # Python CLI runtime dependencies
└── README.md                  # Project documentation
```

---

## 🚀 Getting Started

### Option 1: Browser Web Application (Zero Install)

1. Double-click or open `contact_lot_cleaner.html` in any modern web browser (Chrome, Safari, Edge, Firefox).
2. Drag and drop your `.xlsx` or `.csv` contact spreadsheet.
3. Select your target campaign country.
4. Click **Process & Download Cleaned Lot** to receive your dialer-ready workbook instantly.

### Option 2: Python Terminal CLI

1. **Clone the repository:**
   ```bash
   git clone git@github.com:jgmchavez-22/contact-lot-cleaner.git
   cd contact-lot-cleaner
   ```

2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the script:**
   ```bash
   python contact_lot_cleaner.py
   ```
   The interactive CLI will prompt for the input spreadsheet path and target country.

---

## 🛠️ Tech Stack

- **CLI Engine:** Python 3.9+, [Pandas](https://pandas.pydata.org/), [OpenPyXL](https://openpyxl.readthedocs.io/), [Phonenumbers](https://github.com/daviddrysdale/python-phonenumbers)
- **Web Interface:** HTML5, CSS3, JavaScript (ES6+)
- **Telephony Parsing:** Google's `libphonenumber`
- **Spreadsheet Processing:** SheetJS (`xlsx.full.min.js`)

---

## 🔒 Security & Privacy

All file processing in the web interface occurs **100% locally in your browser session** via JavaScript client runtime. No customer datasets, lead records, or phone numbers are ever transmitted to external servers.
