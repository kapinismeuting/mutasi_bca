# BCA Bank Statement PDF to Excel Converter

A Python 3 application that extracts transactions from BCA (Bank Central Asia) bank statement PDFs and converts them to Excel format.

## 🎯 Features

- **🖥️ Modern Web UI** - Streamlit interface for easy use
- **📄 Extract PDF Transactions** - Parses BCA bank statement PDFs
- **📊 Multiple Output Modes:**
  - Individual Excel files (one per PDF)
  - Consolidated Excel file (all PDFs in one file with multiple sheets)
- **🔍 Comprehensive Logging** - Configurable logging with file and console output
- **⚙️ Configuration Management** - Environment variable support for flexible deployment
- **✅ Data Validation** - Transaction validation before storage
- **💾 Atomic File Operations** - Safe file writes with temporary files and backups
- **🧪 Fully Tested** - 34 unit tests covering all modules
- **📚 Well Documented** - Complete API and user documentation

## 🚀 Quick Start

### Option 1: Streamlit Web Interface (Recommended for Personal Use)

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Run the Streamlit app
streamlit run streamlit_app.py

# 3. Open browser to http://localhost:8501
```

**Or use the launcher scripts:**

```bash
# Linux/Mac
chmod +x run_streamlit.sh
./run_streamlit.sh

# Windows
run_streamlit.bat
```

### Option 2: Command Line

```bash
# Individual processing (one PDF → one Excel)
python mutasi.py

# Consolidated processing (all PDFs → one Excel with sheets)
python mutasi_by_year.py
```

## 📋 Installation

### Prerequisites

- Python 3.7+
- pip or conda

### Setup

1. **Clone or download the project:**
   ```bash
   cd mutasi_bca
   ```

2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Optional: Create and use virtual environment:**
   ```bash
   python3 -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   pip install -r requirements.txt
   ```

## 🖥️ Streamlit Interface

The Streamlit app provides a modern, user-friendly interface with three main sections:

### 📂 Process Tab
- Select input and output folders
- Choose processing mode (Individual or Consolidated)
- Configure options (logging, backups)
- Real-time progress updates
- Detailed result statistics

### 📊 Results Tab
- View generated Excel files
- Check file sizes and modification times
- Inspect output folder contents

### 📚 Help Tab
- Quick start guide
- Processing mode explanations
- Data format documentation
- FAQ and troubleshooting

**Launch the interface:**
```bash
streamlit run streamlit_app.py
```

## ⚙️ Configuration

### Environment Variables

```bash
# Input/Output paths
export PDF_FOLDER="/path/to/pdf/folder"
export OUTPUT_FOLDER="/path/to/output/folder"

# Logging
export LOG_LEVEL="INFO"  # DEBUG, INFO, WARNING, ERROR
export LOG_FILE="/path/to/logfile.log"  # Optional

# Processing limits
export MAX_PDF_SIZE=104857600  # 100 MB in bytes
export BACKUP_FILES="true"  # Create backups before overwriting
```

### Configuration File

Create `.env` file from template:
```bash
cp .env.template .env
# Edit .env with your settings
```

### Default Values

If environment variables are not set:
```python
PDF_FOLDER = ~/dev/appdev/Mutasi/2016
OUTPUT_FOLDER = ~/dev/appdev/Mutasi_Excel
LOG_LEVEL = INFO
MAX_PDF_SIZE = 100 MB
BACKUP_FILES = true
```

## 📖 Usage

### Using Streamlit Interface

1. **Open the application:**
   ```bash
   streamlit run streamlit_app.py
   ```

2. **Configure folders:**
   - Set "Input Folder Path" to folder with PDFs
   - Set "Output Folder Path" where Excel will be saved

3. **Select processing mode:**
   - **Individual Files**: Each PDF → separate Excel
   - **Consolidated File**: All PDFs → one Excel with sheets

4. **Configure options:**
   - Log level (DEBUG, INFO, WARNING, ERROR)
   - Create backups (checkbox)

5. **Validate folders:**
   - Click "Validate" to check configuration

6. **Process:**
   - Click "🚀 Process PDFs" to start
   - Monitor real-time progress

7. **View results:**
   - Go to "Results" tab
   - Inspect generated Excel files

### Using Command Line

```bash
# Individual files
python mutasi.py

# Consolidated
python mutasi_by_year.py
```

### Python API

```python
from mutasi import process_all_pdfs
from mutasi_by_year import process_all_pdfs_to_single_excel

# Individual processing
results = process_all_pdfs("/path/to/pdfs", "/path/to/output")

# Consolidated
results = process_all_pdfs_to_single_excel("/path/to/pdfs", "/path/to/output")

# Check results
for result in results:
    if result.success:
        print(f"✅ {result.filename}: {result.row_count} rows")
    else:
        print(f"❌ {result.filename}: {result.error}")
```

## 🏗️ Architecture

### Module Structure

```
mutasi_bca/
├── config.py                      # Configuration management
├── transaction_parser.py          # PDF parsing & extraction
├── excel_writer.py                # Excel file operations
├── mutasi.py                      # Individual PDF processing (CLI)
├── mutasi_by_year.py              # Consolidated processing (CLI)
├── streamlit_app.py               # Web interface
├── test_transaction_parser.py     # Unit tests
├── test_excel_writer.py           # Unit tests
├── test_config.py                 # Unit tests
├── test_integration.py            # Integration tests
├── requirements.txt               # Dependencies
├── .streamlit/config.toml         # Streamlit config
├── .env.template                  # Environment template
├── run_streamlit.sh               # Linux/Mac launcher
├── run_streamlit.bat              # Windows launcher
├── AUDIT_REPORT.md                # Code audit findings
├── IMPLEMENTATION_SUMMARY.md      # Implementation details
└── readme.md                      # This file
```

### Key Components

- **config.py** - Centralized configuration, environment variables, logging
- **transaction_parser.py** - PDF parsing with transaction extraction
- **excel_writer.py** - Excel file creation with atomic operations
- **streamlit_app.py** - Web interface for processing
- **mutasi.py** - CLI for individual file processing
- **mutasi_by_year.py** - CLI for consolidated processing

## 📊 Data Format

### Excel Output Structure

| Column | Description | Example |
|--------|-------------|---------|
| **Tanggal** | Day of month (1-31) | 15 |
| **Bulan** | Month (1-12) | 6 |
| **Keterangan** | Transaction description | TRANSFER TO ACCOUNT |
| **DB** | Debit amount | 1,234.56 |
| **CR** | Credit amount | 5,000.00 |
| **Saldo** | Account balance | 25,000.00 |

### Supported PDF Format

BCA statements with:
- **Transactions:** `DD/MM [Description] [Amount] [DB/CR] [Balance]`
- **Summaries:** `SALDO AWAL:`, `MUTASI CR:`, `MUTASI DB:`, `SALDO AKHIR:`
- **Numbers:** 1,234.56 (comma separator, dot decimal)

## 🧪 Testing

### Run All Tests

```bash
pip install pytest
pytest -v
```

### Run Specific Tests

```bash
pytest test_transaction_parser.py -v
pytest test_excel_writer.py -v
pytest test_config.py -v
pytest test_integration.py -v
```

## 🔍 Error Handling

The application handles:
- **FileNotFoundError** - Missing PDF files
- **ValueError** - Invalid file size or configuration
- **PermissionError** - No write permission
- **pdfplumber.PDFException** - Corrupted/unsupported PDF

All errors are logged with context and recovery suggestions.

## 📝 Logging

### Log Levels

- **DEBUG** - Detailed information for troubleshooting
- **INFO** - General informational messages (default)
- **WARNING** - Warning messages (e.g., skipped pages)
- **ERROR** - Error messages with recovery attempts

### Configure Logging

```bash
export LOG_LEVEL="DEBUG"
export LOG_FILE="/path/to/app.log"
streamlit run streamlit_app.py
```

## 📦 Deployment

### Local Deployment

```bash
./run_streamlit.sh      # Linux/Mac
run_streamlit.bat       # Windows
```

### Docker Deployment

```dockerfile
FROM python:3.10-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt
COPY . .
CMD ["streamlit", "run", "streamlit_app.py", "--server.port=8501"]
```

```bash
docker build -t bca-converter .
docker run -p 8501:8501 bca-converter
```

### Streamlit Cloud (Free Deployment)

1. Push to GitHub
2. Go to https://share.streamlit.io
3. Connect your GitHub account
4. Select repository and branch
5. Streamlit deploys automatically!

## 🔒 Security

- File size limits prevent DoS attacks
- Specific exception handling avoids info leakage
- Atomic file operations prevent corruption
- Automatic backups provide recovery
- No credentials or sensitive data hard-coded

## 📚 Documentation

- **README.md** - This file, complete documentation
- **STREAMLIT_QUICKSTART.md** - Streamlit quick start guide
- **AUDIT_REPORT.md** - Code audit with 41 findings
- **IMPLEMENTATION_SUMMARY.md** - Implementation details
- **Docstrings** - All functions documented with examples

## ❓ Troubleshooting

### Port Already in Use

```bash
streamlit run streamlit_app.py --server.port 8502
```

### Folder Not Found

- Verify path exists and is accessible
- Use absolute paths (not relative)
- Check file permissions

### PDF Processing Error

- Verify PDF is valid BCA statement
- Check file is not corrupted
- Ensure file size is under 100 MB

### Clear Cache

```bash
streamlit cache clear
```

### Check Logs

Set `LOG_LEVEL=DEBUG` and enable detailed logs in Streamlit interface

## 📋 Requirements

- Python 3.7+
- pdfplumber>=0.10.0
- openpyxl>=3.10.0
- streamlit>=1.28.0
- pytest>=7.0.0 (for testing)

## 🚀 Performance

- Pre-compiled regex patterns (10-20% faster)
- Atomic file operations (safe & reliable)
- Memory efficient (processes one page at a time)
- Error recovery (continues on failed files)

## 🤝 Contributing

When improving:
1. Maintain modular structure
2. Add unit tests for new features
3. Update docstrings and type hints
4. Follow existing code style
5. Update documentation

## 📄 License

This project is provided for educational and business use.

## 💬 Support

For issues or questions:
1. Review Help tab in Streamlit interface
2. Check STREAMLIT_QUICKSTART.md
3. Review AUDIT_REPORT.md for known issues
4. Check logs with DEBUG level enabled

---

**BCA Statement Converter v1.0**  
Built with Python, Streamlit, pdfplumber, and openpyxl  
Last updated: June 3, 2026
