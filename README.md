# ⚡ Legal RPA & OCR Document Parser (Heuristic Engine v9.4)

## 🎯 Business Problem
Managing labor hearings and technical expert inspections (*Insalubridade, Periculosidade, Medical Exams*) required intense manual effort: locating claimant folders across multiple cloud directories, opening scanned PDFs (ERP/SAP employee records, employment contracts, TRCTs), manually copying Job Title, Cost Center, Work Shift, and Hire Date, and creating calendar invites.

## 💡 Technical Solution (End-to-End RPA)
A 100% cloud-native **Robotic Process Automation (RPA)** solution built on Google Apps Script (V8) integrating **Google Sheets, Drive API (OCR), Docs API, and Calendar API**, eliminating external paid licenses while maintaining strict data privacy compliance (LGPD/GDPR).

### 🛠️ Core Engineering Modules
1. **Fuzzy Matching Folder Locator with Scoring Algorithm:** Locates claimant folders across primary and backup repositories by cross-referencing CNJ lawsuit numbers and name variations (stripping accents, prepositions, and typos) using a weighted scoring threshold.
2. **Priority Multi-PDF OCR Pipeline:** Recursively scans up to 10 PDFs per folder, prioritizing high-value HR documents (*Fichas de Registro, FRE, Contracts, TRCT*) and converting scanned images into structured text via Google Drive's native Optical Character Recognition (OCR).
3. **Heuristic Regex Extraction & Sanitization Engine (v9.4):** 4-layer Regex parser calibrated to read misaligned OCR tables from ERP/SAP systems, correcting optical character errors (e.g., `"PRODUGAO"` -> `"PRODUCAO"`), discarding generic placeholders (`"Associado"`), stripping company prefixes from Cost Centers, and normalizing work shifts and hire dates.
4. **Automated Calendar Dispatcher:** Generates structured corporate calendar events and dispatches email invitations to legal counsel and technical experts with 1 click.

## 📂 Repository Structure
* `Code.gs` — Complete RPA & Heuristic Regex Engine (v9.4).
* `appsscript.json` — V8 Runtime manifest and OAuth2 scope configuration.
