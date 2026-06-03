# 🖥️ Streamlit UI Guide

## Overview

The Streamlit interface provides a modern, user-friendly web application for processing BCA bank statement PDFs. This guide walks you through each feature.

---

## 🚀 Launching the App

### Linux/Mac:
```bash
chmod +x run_streamlit.sh
./run_streamlit.sh
```

### Windows:
```bash
run_streamlit.bat
```

### Manual:
```bash
streamlit run streamlit_app.py
```

The app opens at **http://localhost:8501**

---

## 📋 Interface Layout

### Left Sidebar
```
⚙️ Configuration
├── 📋 Processing Mode
│   ├── Individual Files
│   └── Consolidated File
├── 🔍 Logging
│   ├── Log Level (DEBUG/INFO/WARNING/ERROR)
│   └── □ Show detailed logs
├── 🔧 Processing Options
│   └── □ Create backups
└── ℹ️ About
    └── Version info & features
```

### Main Area - Tabs
```
📂 Process | 📊 Results | 📚 Help
```

---

## 📂 TAB 1: Process Files

### Section 1: Input Configuration
```
┌─────────────────────────────────┐
│ 📁 INPUT CONFIGURATION         │
├─────────────────────────────────┤
│                                 │
│ 📁 Input Folder Path:           │
│ ┌─────────────────────────────┐ │
│ │ /path/to/pdf/folder   [🔍] │ │
│ └─────────────────────────────┘ │
│ ⚠️ Folder not found             │
│                                 │
└─────────────────────────────────┘
```

### Section 2: Output Configuration
```
┌─────────────────────────────────┐
│ 📁 OUTPUT CONFIGURATION        │
├─────────────────────────────────┤
│                                 │
│ 📁 Output Folder Path:          │
│ ┌─────────────────────────────┐ │
│ │ /path/to/output       [🔍] │ │
│ └─────────────────────────────┘ │
│ ✅ Output folder ready          │
│                                 │
└─────────────────────────────────┘
```

### Section 3: Start Processing
```
┌─────────────────────────────────┐
│ 🚀 START PROCESSING            │
├─────────────────────────────────┤
│                                 │
│ [🚀 Process PDFs] [✓ Validate]  │
│                                 │
└─────────────────────────────────┘
```

### Section 4a: Validation Results
```
✅ Found 3 PDF file(s)

📋 PDF Files Found
├─ 1. Januari.pdf (2.45 MB)
├─ 2. Februari.pdf (1.89 MB)
└─ 3. Maret.pdf (3.12 MB)

✅ Output folder is valid
```

### Section 4b: Processing Progress
```
🔄 Starting PDF processing...

Progress: [████████████░░░░░░░░] 60%

Processing: Februari.pdf

📊 Summary Statistics
┌────────┬──────────┬────────┬────────┐
│ 3      │ 3        │ 0      │ 58     │
│ Total  │ Success  │ Failed │ Rows   │
└────────┴──────────┴────────┴────────┘

⏱️ Processing completed at 14:30:45

📋 Processing Details

✅ Successful Files
├─ Januari.pdf
│  📁 /path/to/output/Januari.xlsx
│  Rows: 25
├─ Februari.pdf
│  📁 /path/to/output/Februari.xlsx
│  Rows: 18
└─ Maret.pdf
│  📁 /path/to/output/Maret.xlsx
│  Rows: 15
```

---

## 📊 TAB 2: Results

### File Inspection
```
📁 Output Folder to Inspect:
┌─────────────────────────────────┐
│ /path/to/output           [🔍] │
└─────────────────────────────────┘

[🔍 Inspect]

✅ Found 3 Excel file(s)

📋 Output Files
├─ 📄 Januari.xlsx
│  Modified: 2026-06-03 14:30:45
│  Size: 125.34 KB
│  [📋 Copy path]
├─ 📄 Februari.xlsx
│  Modified: 2026-06-03 14:31:12
│  Size: 89.56 KB
│  [📋 Copy path]
└─ 📄 Maret.xlsx
   Modified: 2026-06-03 14:31:45
   Size: 156.78 KB
   [📋 Copy path]
```

---

## 📚 TAB 3: Help

### Quick Start
```
📚 HELP & DOCUMENTATION

🚀 Quick Start
1. Set Input Folder: Choose folder with PDF files
2. Set Output Folder: Choose where to save Excel files
3. Select Mode: Individual or Consolidated
4. Configure Options: Logging level, backups, etc.
5. Click Process: Start converting PDFs
6. View Results: Check output folder for Excel files

💡 Tips
- Validate First: Use "Validate" button before processing
- Monitor Logs: Enable detailed logs for debugging
- Backup Files: Enable automatic backups for safety
- Consolidated Mode: Creates single Excel with multiple sheets
- Individual Mode: Creates separate Excel for each PDF
```

### Processing Modes Comparison
```
📄 Individual Files Mode          📊 Consolidated Mode
┌──────────────────────────────┐  ┌──────────────────────────────┐
│ Input:                       │  │ Input:                       │
│ • mutasi_jan.pdf             │  │ • mutasi_jan.pdf             │
│ • mutasi_feb.pdf             │  │ • mutasi_feb.pdf             │
│                              │  │                              │
│ Output:                      │  │ Output:                      │
│ • mutasi_jan.xlsx            │  │ • 2026.xlsx                  │
│ • mutasi_feb.xlsx            │  │   ├─ JAN (sheet)             │
│                              │  │   └─ FEB (sheet)             │
└──────────────────────────────┘  └──────────────────────────────┘
```

### Data Format
```
📊 Data Format

Column    | Description           | Example
----------|----------------------|----------
Tanggal   | Day of month (1-31)   | 15
Bulan     | Month (1-12)          | 6
Keterangan| Transaction desc      | TRANSFER
DB        | Debit amount          | 1,234.56
CR        | Credit amount         | 5,000.00
Saldo     | Account balance       | 25,000.00
```

### FAQ
```
❓ FAQ

▸ What PDF format is supported?
  BCA bank statements with DD/MM date format

▸ Can I use this with other banks?
  Designed for BCA, can be customized

▸ What if a PDF fails?
  Processor skips and continues with others

▸ How do I use from command line?
  python mutasi.py or python mutasi_by_year.py

▸ Can I schedule automatic processing?
  Yes, with cron (Linux) or Task Scheduler (Windows)
```

---

## 🎨 Sidebar Configuration

### 📋 Processing Mode
```
Select Mode:
( ) Individual Files
(•) Consolidated File
```

**Individual Files:**
- Each PDF → separate Excel file
- Good for organizing by month
- Easy to share individual reports

**Consolidated File:**
- All PDFs → single Excel with sheets
- Good for analysis across months
- Easier to manage one file

### 🔍 Logging
```
Log Level:
┌─────────────────────────────┐
│ INFO                    ▼   │
│ DEBUG                       │
│ WARNING                     │
│ ERROR                       │
└─────────────────────────────┘

☑ Show detailed logs
```

### 🔧 Processing Options
```
☑ Create backups

Creates automatic backups:
file.xlsx → file.xlsx.backup.20260603_143045
```

### ℹ️ About
```
BCA Statement Converter v1.0

Extract transactions from BCA bank statement PDFs
and convert to Excel format.

📊 Processes date, amount, and balance information
💾 Supports backup creation
🔍 Comprehensive error handling
```

---

## 🎯 Common Workflows

### Workflow 1: Individual Monthly Files
```
1. Set Input: ~/Mutasi/2026
2. Set Output: ~/Output
3. Mode: Individual Files
4. Click Validate
5. Click Process
6. Open Results tab
7. Each month's PDF → separate Excel
```

### Workflow 2: Consolidated Annual Report
```
1. Set Input: ~/Mutasi/2026
2. Set Output: ~/Reports
3. Mode: Consolidated File
4. Enable Backups
5. Set Log Level: INFO
6. Click Process
7. Output: 2026.xlsx with 12 sheets (one per month)
```

### Workflow 3: Debug Failed Files
```
1. Enable detailed logs
2. Set Log Level: DEBUG
3. Process files
4. Expand "Failed Files" section
5. View error details
6. Fix issue and reprocess
```

---

## 🔔 Status Indicators

### Input/Output Validation
```
✅ Folder ready          - Valid path, accessible
⚠️  Folder not found     - Path doesn't exist
❌ Invalid path          - Cannot access folder
```

### Processing Status
```
🔄 Processing            - Currently working
⏳ Queued                 - Waiting to process
✅ Successful            - Completed without errors
❌ Failed                - Error occurred
```

### File Operations
```
📄 Excel file found      - Output file exists
📋 Copy path             - Copy file path to clipboard
💾 Backup created        - Automatic backup made
❌ Cannot create backup  - Backup failed (non-critical)
```

---

## ⚙️ Settings Persistence

The Streamlit app remembers:
- Input folder path
- Output folder path
- Selected processing mode
- Log level
- Backup preference

These are stored in Streamlit's session state during the session.

---

## 🚨 Error Messages

### Folder Errors
```
❌ Please enter a valid input folder path
❌ Please enter an output folder path
❌ Folder not found: /invalid/path
```

### File Errors
```
❌ PDF file not found: /path/to/file.pdf
❌ PDF file too large (exceeds 100 MB)
❌ Invalid PDF file: corrupted or unsupported format
```

### Processing Errors
```
❌ Processing Error: [detailed error message]
Details shown in Failed Files section
```

---

## 🎨 Color Scheme

- **Green (#09a339)** - Success, valid
- **Red (#d91e1e)** - Error, failed
- **Orange (#ff9300)** - Warning, skipped
- **Blue (#2E86AB)** - Primary, headers
- **Gray (#F0F2F6)** - Background, secondary

---

## 📱 Responsive Design

The interface works on:
- ✅ Desktop (1920x1080 and higher)
- ✅ Laptop (1366x768 to 1920x1080)
- ⚠️ Tablet (might need horizontal scroll)
- ❌ Mobile (not optimized)

---

## 🔒 Privacy & Security

- ✅ All processing is local
- ✅ No data sent to cloud
- ✅ No file uploads to external servers
- ✅ Configuration files not shared
- ✅ Logs stored locally

---

## 🆘 Troubleshooting

### App Won't Start
```bash
# Check Python version
python --version  # Should be 3.7+

# Check dependencies
pip list | grep streamlit

# Clear Streamlit cache
streamlit cache clear

# Try different port
streamlit run streamlit_app.py --server.port 8502
```

### Browser Won't Open
```
Manual: Open http://localhost:8501 in browser
SSH: Use port forwarding: ssh -L 8501:localhost:8501 user@host
```

### Processing Hangs
```
1. Check input folder has PDF files
2. Check file permissions
3. Enable DEBUG logging
4. Check PDF file is valid
5. Try smaller PDF file first
```

---

## 📞 Support Resources

- **Help Tab** - Built-in documentation
- **STREAMLIT_QUICKSTART.md** - Setup guide
- **README.md** - Full documentation
- **AUDIT_REPORT.md** - Technical details
- **Log Files** - Enable detailed logs for debugging

---

**Last Updated:** June 3, 2026  
**Version:** 1.0
