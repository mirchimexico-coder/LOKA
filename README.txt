# LOKA Restaurant Manager — Bill Auto-Update System
# FREE — No API Key Required
# Last updated: May 2026

## Sheets in LOKA_Restaurant_Manager.xlsx

  1. Capital & Ownership   ← Ownership splits, transfers, pre-opening expenses
  2. Dashboard             ← Live KPIs (auto-updated from Daily Log & Expenses)
  3. Daily Log             ← Enter Date + Revenue + Orders from Soft App daily
  4. Expenses              ← Log every cost (auto-linked to Daily Log)
  5. Inventory             ← Stock levels with low-stock alerts
  6. Staff & Payroll       ← Weekly schedule + auto payroll calculation
  7. Suppliers             ← Purchase orders tracker
  8. Monthly P&L           ← Full P&L Jan–Dec 2026 (formulas auto-calculate)

---

## Two Ways to Log Bills

### Option A — Upload to Chat (Easiest, always free)
1. Take a photo of your bill
2. Upload it directly in your Claude chat
3. Claude reads it and updates your Excel automatically
4. No API key, no setup, no cost

### Option B — Drop in Folder (Fully automatic)
1. Double-click START_WATCHER.bat
2. Drop any bill photo into Bills\_DROP_BILLS_HERE\
3. EasyOCR reads it locally on your PC (no internet)
4. Excel updates and photo is sorted into correct folder
5. A backup is saved to Backups\ automatically (up to 30 versions kept)

---

## Folder Structure

```
Restaurant/
├── LOKA_Restaurant_Manager.xlsx   ← Main Excel (auto-updated)
├── START_WATCHER.bat              ← Double-click to start Option B
├── README.txt                     ← This file
│
├── Backups/                       ← Auto backups (up to 30 versions)
│   └── LOKA_Restaurant_Manager_YYYYMMDD_HHMMSS.xlsx
│
├── Bills/
│   ├── _DROP_BILLS_HERE/          ← DROP PHOTOS HERE (Option B)
│   └── 2026/
│       ├── 05-May/
│       │   └── 2026-05-21/        ← Bills sorted by date automatically
│       ├── 06-June/
│       └── ...
│
├── Scripts/
│   ├── bill_watcher.py            ← Watcher script (EasyOCR)
│   ├── update_capital_sheet.py    ← Refreshes Capital & Ownership sheet
│   └── run_update.bat             ← Double-click to run update_capital_sheet
│
└── Logs/
    └── watcher.log                ← What happened, when
```

---

## Backup System
- Every time a bill is processed (Option B) OR run_update.bat is run,
  a timestamped backup is saved to the Backups\ folder.
- Up to 30 versions are kept. Oldest is deleted when the limit is reached.
- Backup filename format: LOKA_Restaurant_Manager_YYYYMMDD_HHMMSS.xlsx

## Supported Photo Types
JPG, JPEG, PNG, WEBP, BMP

## Tips
- Clear, well-lit photos give best OCR results
- EasyOCR reads Spanish and English text
- If OCR misses something, just upload to Claude chat instead
- All watcher activity is logged in Logs/watcher.log
- First watcher run downloads ~200MB OCR models (one-time only)

## Ownership Summary
  Lohith Reddy           50%   $300,000
  Kashigoud Patil        30%   $180,000
  Shashirekha Basavaraju 20%   $120,000
  TOTAL                 100%   $600,000

  Restaurant purchase: $390,000 (advance $5,000 paid, balance $385,000 pending)
