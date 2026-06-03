# 🚀 BCA Converter - Streamlit Edition - QUICK REFERENCE

## 📦 What You Have Now

✅ **Web Interface** - Modern Streamlit UI  
✅ **CLI Tools** - Command line processing  
✅ **Full Testing** - 34 unit tests  
✅ **Complete Docs** - 5 documentation files  
✅ **Launcher Scripts** - Linux/Mac/Windows  
✅ **Configuration** - Environment-based setup  
✅ **Error Handling** - Comprehensive error management  
✅ **Backups** - Automatic file protection  

---

## 🎯 START HERE

### 1️⃣ Install (One-time)
```bash
pip install -r requirements.txt
```

### 2️⃣ Run (Every time)
```bash
# Linux/Mac
./run_streamlit.sh

# Windows
run_streamlit.bat

# Manual
streamlit run streamlit_app.py
```

### 3️⃣ Use Browser
```
Open: http://localhost:8501
```

### 4️⃣ Process Files
```
1. Set input folder
2. Set output folder
3. Click "Validate"
4. Click "Process"
5. View results
```

---

## 📂 File Overview

### Main Application
```
streamlit_app.py        → Web interface (YOU USE THIS!)
config.py               → Configuration management
transaction_parser.py   → PDF parsing engine
excel_writer.py         → Excel file creation
```

### CLI Tools
```
mutasi.py               → Individual files processing
mutasi_by_year.py       → Consolidated file processing
```

### Tests
```
test_*.py               → 34 unit tests
Run: pytest -v
```

### Documentation
```
README.md               → Full guide
STREAMLIT_INTERFACE.md  → This interface overview
STREAMLIT_UI_GUIDE.md   → Detailed UI walkthrough
STREAMLIT_QUICKSTART.md → Quick start
AUDIT_REPORT.md         → Code quality
```

### Launchers
```
run_streamlit.sh        → Linux/Mac launcher
run_streamlit.bat       → Windows launcher
```

### Config
```
.env.template           → Environment template
requirements.txt        → Python dependencies
.streamlit/config.toml  → Streamlit settings
```

---

## 💡 Common Tasks

### Process Files (Web UI)
```
1. Open Streamlit app
2. Tab: "📂 Process"
3. Enter input folder
4. Enter output folder
5. Click "Process PDFs"
```

### Process from Command Line
```bash
# Individual files (each PDF → separate Excel)
python mutasi.py

# Consolidated (all PDFs → single Excel)
python mutasi_by_year.py
```

### Configure Environment
```bash
export PDF_FOLDER="/path/to/pdfs"
export OUTPUT_FOLDER="/path/to/output"
export LOG_LEVEL="INFO"
streamlit run streamlit_app.py
```

### Run Tests
```bash
pytest -v                           # All tests
pytest test_transaction_parser.py   # Specific test
```

### View Help
Inside Streamlit app → Tab: "📚 Help"

---

## 🎨 Streamlit Interface at a Glance

```
┌─────────────────────────────────────────────────┐
│  ⚙️ Configuration       💰 BCA Converter        │
├──────────────┬──────────────────────────────────┤
│ Processing   │                                  │
│ Mode:        │ 📂 Process | 📊 Results | 📚Help │
│              │                                  │
│ • Individual │ 📁 Input:  [input path]          │
│ • Consol.    │ 📁 Output: [output path]         │
│              │                                  │
│ Log Level:   │ [🚀 Process] [✓ Validate]       │
│ • DEBUG      │                                  │
│ • INFO       │ 📊 Results:                      │
│ • WARNING    │ ✅ 3 files                       │
│ • ERROR      │ ✅ 58 total rows                 │
│              │                                  │
│ ☑ Backups    │ 📋 Details:                      │
│ ☑ Show logs  │ ✅ file1.xlsx                    │
│              │ ✅ file2.xlsx                    │
│              │ ✅ file3.xlsx                    │
│              │                                  │
└──────────────┴──────────────────────────────────┘
```

---

## 🎯 Processing Modes

### Individual Files
```
Input: folder/
  ├─ Jan.pdf
  ├─ Feb.pdf

Output: folder/
  ├─ Jan.xlsx (25 rows)
  ├─ Feb.xlsx (18 rows)
```

### Consolidated
```
Input: folder/
  ├─ Jan.pdf
  ├─ Feb.pdf

Output: folder/
  └─ 2026.xlsx
      ├─ JAN (25 rows)
      ├─ FEB (18 rows)
```

---

## ⚙️ Options

| Option | Default | Effect |
|--------|---------|--------|
| **Log Level** | INFO | How verbose the logging |
| **Backups** | ON | Auto-backup before overwrite |
| **Show Logs** | OFF | Display detailed logging |
| **Mode** | Individual | Processing mode |

---

## 📊 Features Checklist

### Interface
- ✅ Modern web UI with 3 tabs
- ✅ Folder path input with validation
- ✅ Real-time progress updates
- ✅ Processing statistics
- ✅ Error details display
- ✅ File inspection panel

### Processing
- ✅ Individual file mode
- ✅ Consolidated mode
- ✅ Error recovery
- ✅ File validation
- ✅ Automatic backups

### Configuration
- ✅ Environment variables
- ✅ Settings persistence
- ✅ Log level control
- ✅ Folder validation
- ✅ Error messages

### Documentation
- ✅ Built-in help
- ✅ Quick start guide
- ✅ FAQ section
- ✅ Data format info
- ✅ Troubleshooting tips

---

## 🔧 Troubleshooting Quick Fixes

| Problem | Fix |
|---------|-----|
| Port in use | `streamlit run streamlit_app.py --server.port 8502` |
| Folder not found | Use absolute path (e.g., `/home/user/folder`) |
| PDF won't process | Check file isn't corrupted, < 100MB |
| No results | Check input folder has PDFs |
| Permission denied | Check write permissions |

---

## 📱 System Requirements

- ✅ Python 3.7+
- ✅ 100MB disk space (for packages)
- ✅ Internet (for pip install, then works offline)
- ✅ Modern browser (Chrome, Firefox, Safari, Edge)
- ✅ Access to PDF and output folders

---

## 🚀 One-Liner Quick Start

```bash
pip install -r requirements.txt && streamlit run streamlit_app.py
```

Then open: **http://localhost:8501**

---

## 📞 Getting Help

1. **Inside App**: Click "📚 Help" tab
2. **Quick Start**: Read STREAMLIT_QUICKSTART.md
3. **Detailed Guide**: Read STREAMLIT_UI_GUIDE.md
4. **Full Docs**: Read README.md
5. **Technical**: Read IMPLEMENTATION_SUMMARY.md

---

## 🎓 Next Steps

### For Personal Use
1. ✅ Run `./run_streamlit.sh` or `run_streamlit.bat`
2. ✅ Set your input and output folders
3. ✅ Start processing!

### For Automation
1. Use `python mutasi.py` in scripts
2. Set environment variables
3. Schedule with cron or Task Scheduler

### For Sharing
1. Deploy to Streamlit Cloud (free)
2. Or use Docker (see README.md)

---

## 📊 Architecture Diagram

```
┌─────────────────────────────────────┐
│   Streamlit Web Interface           │
│  (streamlit_app.py)                 │
└────────────────┬────────────────────┘
                 │
         ┌───────┴────────┐
         │                │
         ▼                ▼
┌──────────────────┐  ┌──────────────────┐
│ config.py        │  │ CLI Tools        │
│ (Configuration)  │  │ (mutasi.py)      │
└──────────────────┘  └──────────────────┘
         │                │
         └───────┬────────┘
                 │
         ┌───────┴────────┐
         │                │
         ▼                ▼
┌──────────────────┐  ┌──────────────────┐
│ transaction_     │  │ excel_writer.py  │
│ parser.py        │  │ (Excel Creation) │
└──────────────────┘  └──────────────────┘
         │                │
         └───────┬────────┘
                 │
         ┌───────▼────────┐
         │ Input/Output   │
         │ (PDF & Excel)  │
         └────────────────┘
```

---

## ✨ Key Improvements Over Command Line

| Feature | CLI | Web UI |
|---------|-----|--------|
| Visual Feedback | ❌ | ✅ |
| Progress Tracking | ❌ | ✅ |
| File Browser | ❌ | ✅ |
| Result Preview | ❌ | ✅ |
| Settings UI | ❌ | ✅ |
| Help System | ❌ | ✅ |
| Ease of Use | Medium | Easy |

---

## 🎉 Summary

You now have a **production-ready, professional web interface** for your BCA converter:

✅ Modern, intuitive UI  
✅ Complete documentation  
✅ Full testing suite  
✅ Easy setup and launch  
✅ Powerful configuration  
✅ Reliable error handling  

**Ready to use? Just run:**
```bash
./run_streamlit.sh
```

Then open **http://localhost:8501** 🚀

---

**Version:** 1.0  
**Status:** ✅ Ready to Use  
**Date:** June 3, 2026
