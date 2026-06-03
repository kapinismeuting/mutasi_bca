#!/usr/bin/env python3
"""
Streamlit UI for BCA Bank Statement PDF to Excel Converter.
Modern web interface for processing bank statements.
"""

import streamlit as st
import os
import shutil
import zipfile
import io
from pathlib import Path
from datetime import datetime
from config import Config, get_logger
from mutasi import process_all_pdfs, ProcessResult
from mutasi_by_year import process_all_pdfs_to_single_excel

logger = get_logger('streamlit_app')

# ============================================================================
# PAGE CONFIGURATION
# ============================================================================

st.set_page_config(
    page_title="BCA Statement Converter",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
    <style>
    .main {
        padding-top: 2rem;
    }
    .stTabs [data-baseweb="tab-list"] button {
        font-size: 16px;
        padding: 10px 20px;
    }
    .metric-card {
        background-color: #f0f2f6;
        padding: 1.5rem;
        border-radius: 0.5rem;
        margin: 0.5rem 0;
    }
    .success-text {
        color: #09a339;
        font-weight: bold;
    }
    .error-text {
        color: #d91e1e;
        font-weight: bold;
    }
    .warning-text {
        color: #ff9300;
        font-weight: bold;
    }
    </style>
    """, unsafe_allow_html=True)

# ============================================================================
# SIDEBAR CONFIGURATION
# ============================================================================

st.sidebar.title("⚙️ Configuration")

with st.sidebar:
    st.divider()
    
    # About
    st.subheader("ℹ️ About")
    st.info(
        "**BCA Statement Converter v1.0**\n\n"
        "Extract transactions from BCA bank statement PDFs "
        "and convert to Excel format.\n\n"
        "📤 Upload PDFs directly\n"
        "📥 Download Excel results\n"
        "🔍 Multiple processing modes"
    )
    
    st.divider()
    
    st.subheader("🎯 Quick Tips")
    st.markdown("""
    1. **Upload PDFs** on the Process tab
    2. **Select mode**: Individual or Consolidated
    3. **Click Process** to convert
    4. **Download files** from results section
    
    Individual mode creates separate Excel files per PDF.
    
    Consolidated mode combines all PDFs into one Excel with sheets.
    """)

# ============================================================================
# MAIN PAGE
# ============================================================================

st.title("💰 BCA Bank Statement Converter")
st.write("Extract transactions from BCA bank statements and convert to Excel")

# Create tabs for different sections
tab1, tab2, tab3 = st.tabs(["📂 Process", "📊 Results", "📚 Help"])

# ============================================================================
# TAB 1: PROCESS
# ============================================================================

with tab1:
    st.subheader("📤 Upload PDF Files")
    st.write("Upload your BCA bank statement PDF files below")
    
    # File uploader
    uploaded_files = st.file_uploader(
        "Choose PDF files",
        type="pdf",
        accept_multiple_files=True,
        help="Select one or more PDF files to process"
    )
    
    if uploaded_files:
        st.success(f"✅ {len(uploaded_files)} file(s) selected")
        
        # Show file list
        with st.expander("📋 Selected Files"):
            for idx, file in enumerate(uploaded_files, 1):
                file_size = len(file.getvalue()) / (1024 * 1024)
                st.write(f"{idx}. {file.name} ({file_size:.2f} MB)")
    
    st.divider()
    
    # Configuration
    st.subheader("⚙️ Processing Configuration")
    
    col1, col2 = st.columns(2)
    
    with col1:
        processing_mode = st.radio(
            "Processing Mode:",
            ["Individual Files", "Consolidated File"],
            help="Individual: Each PDF → Separate Excel\n\n"
                 "Consolidated: All PDFs → Single Excel with sheets",
            key="process_mode"
        )
    
    with col2:
        create_backups = st.checkbox(
            "💾 Create backups",
            value=True,
            help="Create automatic backups before overwriting files"
        )
    
    st.divider()
    
    # Process button
    process_button = st.button(
        "🚀 Process PDFs",
        type="primary",
        use_container_width=True,
        disabled=not uploaded_files,
        help="Start processing the uploaded PDF files"
    )
    
    # Processing
    if process_button and uploaded_files:
        import tempfile
        try:
            with tempfile.TemporaryDirectory() as input_dir, tempfile.TemporaryDirectory() as output_dir:
                temp_input_dir = Path(input_dir)
                temp_output_dir = Path(output_dir)
                
                # Save uploaded files to temp directory
                with st.status("Preparing files...", expanded=True):
                    st.write("Saving uploaded PDF files...")
                    
                    for uploaded_file in uploaded_files:
                        if uploaded_file.size > Config.MAX_PDF_SIZE:
                            st.error(f"File {uploaded_file.name} exceeds max size limit.")
                            st.stop()
                        
                        file_path = temp_input_dir / uploaded_file.name
                        file_path.write_bytes(uploaded_file.getvalue())
                        st.write(f"✅ Saved: {uploaded_file.name}")
            
                # Process PDFs
                with st.status("Processing PDF files...", expanded=True):
                    st.write(f"🔄 Starting {processing_mode.lower()} mode...")
                    
                    # Process based on mode
                    if processing_mode == "Individual Files":
                        results = process_all_pdfs(str(temp_input_dir), str(temp_output_dir), backup_files=create_backups)
                    else:
                        results = process_all_pdfs_to_single_excel(
                            str(temp_input_dir), str(temp_output_dir), backup_files=create_backups)
                    
                    st.write("✅ Processing completed!")
            
                # Calculate statistics
                successful = sum(1 for r in results if r.success)
                failed = sum(1 for r in results if not r.success)
                total_rows = sum(r.row_count for r in results if r.success)
                
                # Display results
                st.divider()
                st.success("✅ Processing Complete!")
            
                # Metrics
                st.subheader("📊 Summary Statistics")
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    st.metric("Total Files", len(results))
                with col2:
                    st.metric("Successful", successful, delta=None, delta_color="off")
                with col3:
                    st.metric("Failed", failed, delta=None, delta_color="off")
                with col4:
                    st.metric("Total Rows", total_rows)
                
                st.info(f"⏱️ Processing completed at {datetime.now().strftime('%H:%M:%S')}")
                
                # Detailed results
                st.subheader("📋 Processing Details")
                
                # Get output files
                output_files = list(temp_output_dir.glob("*.xlsx"))
            
                if output_files:
                    st.success(f"✅ Generated {len(output_files)} Excel file(s)")
                    
                    with st.expander("📥 Download Files", expanded=True):
                        st.write("Click on a file to download:")
                        
                        # Individual downloads
                        col1, col2, col3 = st.columns(3)
                        
                        for idx, file_path in enumerate(sorted(output_files)):
                            col = [col1, col2, col3][idx % 3]
                            
                            with col:
                                file_data = file_path.read_bytes()
                                st.download_button(
                                    label=f"📥 {file_path.name}",
                                    data=file_data,
                                    file_name=file_path.name,
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True
                                )
                        
                        st.divider()
                        
                        # Download all as zip
                        if len(output_files) > 1:
                            st.write("Or download all files as ZIP:")
                            
                            zip_buffer = io.BytesIO()
                            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                                for file_path in sorted(output_files):
                                    zip_file.write(file_path, arcname=file_path.name)
                            
                            zip_buffer.seek(0)
                            
                            st.download_button(
                                label="📦 Download All as ZIP",
                                data=zip_buffer.getvalue(),
                                file_name=f"bca_statements_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                                mime="application/zip",
                                use_container_width=True
                            )
                
                # Successful files details
                if successful > 0:
                    with st.expander("✅ Successful Files", expanded=False):
                        for result in results:
                            if result.success:
                                col1, col2 = st.columns([3, 1])
                                with col1:
                                    st.write(f"**{result.filename}**")
                                with col2:
                                    st.metric("Rows", result.row_count, label_visibility="collapsed")
            
                # Failed files details
                if failed > 0:
                    with st.expander("❌ Failed Files", expanded=True):
                        for result in results:
                            if not result.success:
                                st.error(f"**{result.filename}**")
                                st.code(result.error, language="text")
        
        except Exception as e:
            st.error(f"❌ Processing Error: {str(e)}")
            logger.exception(f"Processing error: {e}")
    
    elif process_button and not uploaded_files:
        st.warning("⚠️ Please upload at least one PDF file")


# ============================================================================
# TAB 2: RESULTS
# ============================================================================

with tab2:
    st.subheader("📊 Processing Results")
    
    st.info("""
    ℹ️ **About this tab:**
    
    After processing PDF files in the **Process** tab, you can download the 
    generated Excel files directly from the results section.
    
    The download buttons appear automatically after processing completes.
    """)
    
    st.divider()
    
    st.subheader("📁 Manual File Inspection (Optional)")
    st.write("If you process files from your file system, inspect the output folder here:")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.info(f"Directory locked to: {Config.OUTPUT_FOLDER}")
        output_path = Config.OUTPUT_FOLDER
    
    with col2:
        inspect_button = st.button("🔍 Inspect Folder", use_container_width=True)
    
    if inspect_button and output_path:
        if os.path.isdir(output_path):
            excel_files = [f for f in os.listdir(output_path) 
                          if f.lower().endswith('.xlsx')]
            
            if excel_files:
                st.success(f"✅ Found {len(excel_files)} Excel file(s)")
                
                st.subheader("📋 Output Files")
                
                for filename in sorted(excel_files):
                    file_path = os.path.join(output_path, filename)
                    file_size = os.path.getsize(file_path) / (1024 * 1024)
                    file_time = datetime.fromtimestamp(
                        os.path.getmtime(file_path)
                    ).strftime("%Y-%m-%d %H:%M:%S")
                    
                    col1, col2, col3 = st.columns([3, 1, 1])
                    
                    with col1:
                        st.write(f"📄 **{filename}**")
                        st.caption(f"Modified: {file_time}")
                    with col2:
                        st.metric("Size", f"{file_size:.2f} MB", label_visibility="collapsed")
                    with col3:
                        if st.button("📋", key=f"copy_{filename}", help="Copy path"):
                            st.info(f"Path: {file_path}")
            else:
                st.warning("⚠️ No Excel files found in output folder")
        else:
            st.error("❌ Output folder not found")

# ============================================================================
# TAB 3: HELP
# ============================================================================

with tab3:
    st.subheader("📚 Help & Documentation")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("🚀 Quick Start")
        st.markdown("""
        1. **Set Input Folder**: Choose folder with PDF files
        2. **Set Output Folder**: Choose where to save Excel files
        3. **Select Mode**: Individual or Consolidated
        4. **Configure Options**: Logging level, backups, etc.
        5. **Click Process**: Start converting PDFs
        6. **View Results**: Check output folder for Excel files
        """)
    
    with col2:
        st.subheader("💡 Tips")
        st.markdown("""
        - **Validate First**: Use "Validate" button before processing
        - **Monitor Logs**: Enable detailed logs for debugging
        - **Backup Files**: Enable automatic backups for safety
        - **Consolidated Mode**: Creates single Excel with multiple sheets
        - **Individual Mode**: Creates separate Excel for each PDF
        """)
    
    st.divider()
    
    st.subheader("📋 Processing Modes")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.info("""
        **📄 Individual Files Mode**
        
        Processes each PDF into a separate Excel file.
        
        Input:
        - mutasi_jan.pdf
        - mutasi_feb.pdf
        
        Output:
        - mutasi_jan.xlsx
        - mutasi_feb.xlsx
        """)
    
    with col2:
        st.info("""
        **📊 Consolidated Mode**
        
        Combines all PDFs into a single Excel file with 
        one sheet per month.
        
        Input:
        - mutasi_jan.pdf
        - mutasi_feb.pdf
        
        Output:
        - 2026.xlsx (with JAN, FEB sheets)
        """)
    
    st.divider()
    
    st.subheader("📊 Data Format")
    st.markdown("""
    Each transaction row contains:
    
    | Column | Description | Example |
    |--------|-------------|---------|
    | **Tanggal** | Day of month (1-31) | 15 |
    | **Bulan** | Month (1-12) | 6 |
    | **Keterangan** | Transaction description | TRANSFER TO ACCOUNT |
    | **DB** | Debit amount | 1,234.56 |
    | **CR** | Credit amount | 5,000.00 |
    | **Saldo** | Account balance | 25,000.00 |
    """)
    
    st.divider()
    
    st.subheader("❓ FAQ")
    
    with st.expander("What PDF format is supported?"):
        st.write("""
        BCA bank statements with:
        - Transaction rows: `DD/MM [Description] [Amount] [DB/CR] [Balance]`
        - Summary rows: `SALDO AWAL:`, `MUTASI CR:`, `MUTASI DB:`, `SALDO AKHIR:`
        - Number format: 1,234.56 (comma separator, dot decimal)
        """)
    
    with st.expander("Can I use this with other banks?"):
        st.write("""
        Currently designed for BCA format. To support other banks:
        1. Update regex patterns in config
        2. Adjust parsing logic for different formats
        3. Contact support for customization
        """)
    
    with st.expander("What if a PDF fails to process?"):
        st.write("""
        The processor will:
        1. Skip the failed PDF and continue with others
        2. Log the error with details
        3. Report the error in the Results section
        4. Show which files succeeded and which failed
        """)
    
    with st.expander("How do I use this from command line?"):
        st.write("""
        ```bash
        # Run Streamlit app
        streamlit run streamlit_app.py
        
        # Or use Python CLI
        export PDF_FOLDER="/path/to/pdfs"
        export OUTPUT_FOLDER="/path/to/output"
        python mutasi.py
        ```
        """)
    
    with st.expander("Can I schedule this to run automatically?"):
        st.write("""
        Yes! Use:
        - **Linux/Mac**: cron jobs
        - **Windows**: Task Scheduler
        - **Docker**: Container scheduling
        
        Example cron (daily at 2 AM):
        ```
        0 2 * * * cd /path && python mutasi.py
        ```
        """)

# ============================================================================
# FOOTER
# ============================================================================

st.divider()

col1, col2, col3 = st.columns(3)

with col1:
    st.caption("💰 BCA Statement Converter v1.0")

with col2:
    st.caption(f"📅 {datetime.now().strftime('%Y-%m-%d')}")

with col3:
    st.caption("🔐 Processing local files securely")
