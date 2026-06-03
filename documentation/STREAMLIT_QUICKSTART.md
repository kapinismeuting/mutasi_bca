#!/usr/bin/env python3
"""
BCA Statement Converter - Quick Start Guide
============================================

This file contains setup and running instructions.
"""

"""
# 🚀 QUICK START

## 1. Install Dependencies

```bash
pip install -r requirements.txt
```

## 2. Run the Streamlit App

```bash
streamlit run streamlit_app.py
```

The app will open in your default browser at:
http://localhost:8501

## 3. Using the App

### Setup:
1. In the left sidebar, select your processing mode (Individual or Consolidated)
2. Enter your PDF input folder path
3. Enter your output folder path
4. Configure options (backups, log level, etc.)

### Process:
1. Click "Validate" to check your folders
2. Click "🚀 Process PDFs" to start processing
3. Monitor progress in real-time
4. View results with detailed statistics

### View Results:
1. Go to "Results" tab
2. Enter output folder path
3. Click "Inspect" to see generated Excel files

## 4. Environment Setup (Optional)

Set default folders via environment variables:

```bash
export PDF_FOLDER="/path/to/your/pdfs"
export OUTPUT_FOLDER="/path/to/your/output"
export LOG_LEVEL="INFO"
```

Then run:
```bash
streamlit run streamlit_app.py
```

## 5. Docker Setup (Optional)

Create Dockerfile:

```dockerfile
FROM python:3.10-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt
COPY . .
CMD ["streamlit", "run", "streamlit_app.py"]
```

Run:
```bash
docker build -t bca-converter .
docker run -p 8501:8501 bca-converter
```

# 📊 Features

✅ Modern web interface
✅ Real-time progress updates
✅ File validation
✅ Detailed error reporting
✅ Processing statistics
✅ Two processing modes:
   - Individual: Each PDF → Separate Excel
   - Consolidated: All PDFs → Single Excel with sheets
✅ Configurable logging
✅ Automatic backups

# 🎯 Use Cases

1. **Personal Use**: Run locally for batch processing
2. **Daily Processing**: Schedule with cron/Task Scheduler
3. **Cloud Deployment**: Deploy to Streamlit Cloud (free)
4. **Server Setup**: Run on dedicated server for team access

# 🔧 Troubleshooting

## Port Already in Use
```bash
streamlit run streamlit_app.py --server.port 8502
```

## Clear Cache
```bash
streamlit cache clear
```

## Check Logs
Look at `.streamlit/logs/` directory

## Reset Configuration
```bash
rm -rf ~/.streamlit/
```

# 📚 Command Line Alternative

If you prefer command line:

```bash
# Individual processing
python mutasi.py

# Consolidated processing
python mutasi_by_year.py
```

# 🌐 Deploy to Streamlit Cloud

1. Push code to GitHub
2. Go to https://share.streamlit.io
3. Connect your GitHub account
4. Select repository and branch
5. Streamlit deploys automatically!

# 📞 Support

For issues:
1. Check logs in left sidebar (🔍 Show detailed logs)
2. Review Help tab in the app
3. Check AUDIT_REPORT.md for known issues
4. Check IMPLEMENTATION_SUMMARY.md for architecture

# 📖 Documentation

- README.md - Full documentation
- AUDIT_REPORT.md - Code audit findings
- IMPLEMENTATION_SUMMARY.md - Implementation details
- This file - Quick start guide

"""

if __name__ == "__main__":
    print(__doc__)
