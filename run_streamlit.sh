#!/bin/bash
# BCA Statement Converter - Streamlit Launcher
# Simple script to start the Streamlit application

set -e

# Colors for output
GREEN='\033[0;32m'
BLUE='\033[0;34m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# Print header
echo -e "${BLUE}╔════════════════════════════════════════════╗${NC}"
echo -e "${BLUE}║   BCA Statement Converter - Streamlit      ║${NC}"
echo -e "${BLUE}╚════════════════════════════════════════════╝${NC}"
echo ""

# Check Python
if ! command -v python3 &> /dev/null; then
    echo -e "${YELLOW}❌ Python 3 not found. Please install Python 3.7+${NC}"
    exit 1
fi
echo -e "${GREEN}✓ Python found:${NC} $(python3 --version)"

# Check if virtual environment is needed
if [ ! -d "venv" ]; then
    echo -e "${YELLOW}Creating virtual environment...${NC}"
    python3 -m venv venv
fi

# Activate virtual environment
source venv/bin/activate 2>/dev/null || . venv/Scripts/activate 2>/dev/null
echo -e "${GREEN}✓ Virtual environment activated${NC}"

# Install/upgrade requirements
echo -e "${YELLOW}Checking dependencies...${NC}"
pip install -q --upgrade pip
pip install -q -r requirements.txt
echo -e "${GREEN}✓ Dependencies installed${NC}"

# Load environment variables if .env exists
if [ -f ".env" ]; then
    echo -e "${GREEN}✓ Loading environment variables from .env${NC}"
    export $(cat .env | grep -v '^#' | xargs)
fi

# Show configuration
echo ""
echo -e "${BLUE}Configuration:${NC}"
echo -e "  PDF Folder:    ${PDF_FOLDER:-${HOME}/dev/appdev/Mutasi/2016}"
echo -e "  Output Folder: ${OUTPUT_FOLDER:-${HOME}/dev/appdev/Mutasi_Excel}"
echo -e "  Log Level:     ${LOG_LEVEL:-INFO}"
echo ""

# Start Streamlit
echo -e "${GREEN}🚀 Starting Streamlit application...${NC}"
echo -e "${YELLOW}Opening browser...${NC}"
echo -e "${BLUE}If browser doesn't open, visit: http://localhost:8501${NC}"
echo ""

streamlit run streamlit_app.py

# Deactivate on exit
deactivate 2>/dev/null || true
