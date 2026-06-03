# Comprehensive Codebase Audit Report: BCA Statement Extractor

## 1. Architecture Issues

### Global Mutable State in Web Application
- **Severity**: High
- **File Location**: `streamlit_app.py` (Line 194), `config.py`
- **Description**: The application modifies global state (`Config.BACKUP_FILES = create_backups`) during request processing. 
- **Why it is a problem**: Streamlit runs a persistent process where multiple users/sessions share the same memory space. Mutating global configuration based on a single user's session will affect all other concurrent users, leading to unpredictable behavior and race conditions.
- **Recommended Fix**: Remove global mutable configuration. Pass session-specific configurations (like `backup_files=True/False`) directly to the processing functions (`process_all_pdfs`, etc.) as arguments.

### Tight Coupling of UI and Business Logic
- **Severity**: Medium
- **File Location**: `streamlit_app.py`
- **Description**: The Streamlit interface contains raw file system operations, zip file generation, and temporary directory management.
- **Why it is a problem**: Violates the Single Responsibility Principle, making the UI hard to test and maintain. It also limits reusability of these features in other interfaces (like a CLI or REST API).
- **Recommended Fix**: Extract file management, zip creation, and temp directory handling into a dedicated `file_manager.py` or `services.py` module.

---

## 2. Potential Bugs

### Broken Test Case / Logic Mismatch
- **Severity**: Medium
- **File Location**: `test_transaction_parser.py` (Line 86), `transaction_parser.py` (Line 88)
- **Description**: `test_extract_date_no_leading_zeros` expects `extract_date(" 15/06 Transfer")` to return `None`. However, `extract_date` calls `line.strip()` before running the regex, meaning it *will* match and return `(15, 6)`.
- **Why it is a problem**: The test is broken and misrepresents the actual behavior of the system. This indicates either a flaw in the regex design or a faulty test.
- **Recommended Fix**: If leading spaces should invalidate a match, remove `.strip()` from the regex matching step, or update the test to reflect the intended behavior (that leading spaces are stripped and allowed).

### Unhandled Temp File Cleanup (Resource Leak)
- **Severity**: Medium
- **File Location**: `streamlit_app.py` (Lines 302-303)
- **Description**: The `finally` block in the processing flow has a `pass` statement, assuming Streamlit cleans up temp directories when the browser session ends.
- **Why it is a problem**: Streamlit does not automatically delete arbitrary directories created in `/tmp`. This will lead to disk space exhaustion over time as users upload PDFs.
- **Recommended Fix**: Use Python's `tempfile.TemporaryDirectory()` which automatically cleans up upon garbage collection/context exit, or explicitly run `shutil.rmtree()` in the `finally` block.

### Unsafe Variable Reference in Exception Block
- **Severity**: Low
- **File Location**: `excel_writer.py` (Lines 116-118)
- **Description**: `if os.path.exists(temp_path):` is in the `except` block, but if the exception occurs before `temp_path` is assigned, this will raise an `UnboundLocalError`.
- **Why it is a problem**: Masks the original exception with a new, confusing error.
- **Recommended Fix**: Initialize `temp_path = None` at the top of the `try` block and check `if temp_path and os.path.exists(temp_path):` in the `except` block.

---

## 3. Security Vulnerabilities

### Path Traversal / Arbitrary File Read
- **Severity**: Critical
- **File Location**: `streamlit_app.py` (Lines 333-368)
- **Description**: The manual file inspection feature allows users to input any directory path (`output_path = st.text_input(...)`), and the app will execute `os.listdir(output_path)` and expose file metadata and contents (via copy path).
- **Why it is a problem**: A malicious user could input sensitive paths (e.g., `/etc/`, `/var/log`, or `C:\Windows`) and enumerate server files, potentially exposing sensitive data.
- **Recommended Fix**: Restrict the directory listing to only a predefined safe directory (e.g., the app's designated output folder). Validate that the resolved absolute path of the user's input is a subdirectory of the safe directory.

### Denial of Service (DoS) via Unrestricted Uploads
- **Severity**: High
- **File Location**: `streamlit_app.py` (Lines 185-186)
- **Description**: Uploaded files are written directly to memory (`uploaded_file.getvalue()`) and then to disk without checking file sizes *before* saving. 
- **Why it is a problem**: A user could upload extremely large files or thousands of files, causing out-of-memory (OOM) errors or disk exhaustion on the server. The `MAX_PDF_SIZE` in `config.py` is only checked *after* the file is parsed by `pdfplumber`, which is too late.
- **Recommended Fix**: Enforce Streamlit's built-in file upload size limit (`server.maxUploadSize`) and validate file sizes before reading them into memory.

---

## 4. Performance Bottlenecks

### In-Memory File Duplication
- **Severity**: Medium
- **File Location**: `streamlit_app.py` (Line 186)
- **Description**: `file_path.write_bytes(uploaded_file.getvalue())` loads the entire uploaded file into RAM as a bytes object.
- **Why it is a problem**: For large PDFs or many concurrent users, holding full file contents in memory will cause high RAM usage.
- **Recommended Fix**: Stream the uploaded file to disk in chunks instead of loading the entire content into memory at once.

### PDF Parsing Overhead
- **Severity**: Medium
- **File Location**: `transaction_parser.py` (Lines 266-301)
- **Description**: Iterating through all pages and extracting text block by block is synchronous and computationally expensive.
- **Why it is a problem**: Large PDFs will cause the Streamlit app to hang and could lead to timeouts.
- **Recommended Fix**: Implement multiprocessing or async processing for PDF parsing. Additionally, use caching where possible if the same PDF is uploaded multiple times.

---

## 5. Database Risks

### Concurrency Issues with File Storage
- **Severity**: Medium
- **File Location**: `excel_writer.py` (Lines 95-112)
- **Description**: The application uses a flat-file storage system for outputs. While atomic renames (`os.replace`) are used, concurrent requests writing to the exact same file name will cause race conditions.
- **Why it is a problem**: In `streamlit_app.py`, files are saved using the original uploaded filename. If two users upload `mutasi.pdf` simultaneously, they will overwrite each other's files.
- **Recommended Fix**: Append a UUID or timestamp to filenames internally to guarantee uniqueness. 

---

## 6. API Design Issues

### Inconsistent Data Structures
- **Severity**: Medium
- **File Location**: `transaction_parser.py`
- **Description**: `parse_transaction_block` yields a `Transaction` object, but `parse_summary_line` yields a raw `Dict`. The main function `parse_bca_transactions` converts everything to `Dict` before returning.
- **Why it is a problem**: Lack of uniform data types makes the downstream logic (like `excel_writer.py`) fragile and reliant on string keys (`tx.get('tanggal')`), defeating the purpose of having a `Transaction` class.
- **Recommended Fix**: Create a base class or `TypedDict` for all rows (both transactions and summaries), and ensure `parse_bca_transactions` returns a list of strongly typed objects.

### String-Based Monetary Values
- **Severity**: Medium
- **File Location**: `transaction_parser.py` (Lines 110-137)
- **Description**: Debit, credit, and balance amounts are stored and passed around as strings (e.g., `"1,000.00"`).
- **Why it is a problem**: Excel treats these as text unless specifically parsed, which breaks spreadsheet calculations for the user. It also makes mathematical validation in Python impossible.
- **Recommended Fix**: Parse strings into Python `Decimal` or `float` objects during extraction, and write them to Excel as numeric types.

---

## 7. Code Smells

### Broad Exception Catching
- **Severity**: Medium
- **File Location**: `mutasi.py` (Line 78), `excel_writer.py` (Lines 166, 226)
- **Description**: Extensive use of `except Exception as e:`.
- **Why it is a problem**: Masks critical system errors (like `KeyboardInterrupt`, `MemoryError`, or syntax errors in development) and makes debugging difficult.
- **Recommended Fix**: Catch specific exceptions (`IOError`, `ValueError`, `pdfplumber.PDFException`), and let unexpected exceptions bubble up.

### Code Duplication
- **Severity**: Low
- **File Location**: `mutasi.py` vs `mutasi_by_year.py`
- **Description**: The core logic for iterating over folders, validating input, and tracking successes/failures is duplicated across both files.
- **Why it is a problem**: Increases maintenance burden; bug fixes must be applied in two places.
- **Recommended Fix**: Abstract the directory traversal and file handling logic into a single generic function that accepts a processing callback.

---

## 8. Dead Code

### Misplaced / Redundant Imports
- **Severity**: Low
- **File Location**: `transaction_parser.py` (Line 250)
- **Description**: `import os` is declared inside the `parse_bca_transactions` function.
- **Why it is a problem**: It violates PEP 8 standard conventions and adds a tiny (though negligible) overhead.
- **Recommended Fix**: Move `import os` to the top of the file.

---

## 9. Missing Tests

### Critical Path Untested
- **Severity**: High
- **File Location**: `test_transaction_parser.py`
- **Description**: While small utility functions (`extract_date`, etc.) are tested, the core parsing loop `parse_transaction_block` and the main PDF extraction `parse_bca_transactions` have zero test coverage.
- **Why it is a problem**: Any changes to the layout parsing or PDF text extraction logic could break the application without the test suite catching it.
- **Recommended Fix**: Add unit tests using mock PDF text blocks for `parse_transaction_block`. Add integration tests using sample/dummy PDFs for `parse_bca_transactions`.

### UI and API E2E Tests
- **Severity**: High
- **File Location**: `streamlit_app.py`
- **Description**: There are no tests for the Streamlit UI or the end-to-end integration of file uploads to Excel downloads.
- **Why it is a problem**: UI regressions or broken session state logic will go unnoticed until manually tested.
- **Recommended Fix**: Implement Streamlit AppTest framework tests to simulate user uploads and verify UI behavior.
