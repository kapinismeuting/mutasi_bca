#!/usr/bin/env python3
"""
Extract BCA bank statement transactions to Excel.
Processes each PDF into individual Excel files.

Configuration via environment variables:
- PDF_FOLDER: Input directory with PDFs
- OUTPUT_FOLDER: Output directory for Excel files
- LOG_LEVEL: Logging level (DEBUG, INFO, WARNING, ERROR)
- LOG_FILE: Optional log file path
"""

import os
import sys
from pathlib import Path
from typing import List, Optional
from dataclasses import dataclass

from config import Config, get_logger
from transaction_parser import parse_bca_transactions
from excel_writer import write_transactions_to_excel

logger = get_logger('mutasi.main')


@dataclass
class ProcessResult:
    """Result of processing a single PDF."""
    success: bool
    filename: str
    output_path: Optional[str] = None
    error: Optional[str] = None
    row_count: int = 0


def process_single_pdf(pdf_path: str, output_folder: str) -> ProcessResult:
    """Process a single PDF file and write to Excel.
    
    Args:
        pdf_path: Path to PDF file
        output_folder: Output directory for Excel file
        
    Returns:
        ProcessResult with success status and details
    """
    filename = os.path.basename(pdf_path)
    
    try:
        logger.info(f"Processing: {filename}")
        
        # Parse transactions
        transactions = parse_bca_transactions(pdf_path)
        
        # Generate output path
        base_name = os.path.splitext(filename)[0]
        output_path = os.path.join(output_folder, f"{base_name}.xlsx")
        
        # Write to Excel
        write_transactions_to_excel(transactions, output_path)
        
        return ProcessResult(
            success=True,
            filename=filename,
            output_path=output_path,
            row_count=len(transactions)
        )
        
    except FileNotFoundError as e:
        error_msg = f"PDF file not found: {e}"
        logger.error(error_msg)
        return ProcessResult(success=False, filename=filename, error=error_msg)
    
    except ValueError as e:
        error_msg = f"Invalid input: {e}"
        logger.error(error_msg)
        return ProcessResult(success=False, filename=filename, error=error_msg)
    
    except Exception as e:
        error_msg = f"Error: {e}"
        logger.error(error_msg)
        return ProcessResult(success=False, filename=filename, error=error_msg)


def process_all_pdfs(pdf_folder: str, output_folder: str) -> List[ProcessResult]:
    """Process all PDFs in folder.
    
    Args:
        pdf_folder: Input directory with PDFs
        output_folder: Output directory for Excel files
        
    Returns:
        List of ProcessResult for each file
    """
    results = []
    
    # Validate input folder
    if not os.path.isdir(pdf_folder):
        logger.error(f"Folder not found: {pdf_folder}")
        raise ValueError(f"Folder not found: {pdf_folder}")
    
    # Find PDF files
    pdf_files = sorted([
        f for f in os.listdir(pdf_folder) 
        if f.lower().endswith('.pdf')
    ])
    
    if not pdf_files:
        logger.warning(f"No PDFs found in {pdf_folder}")
        return results
    
    logger.info(f"Found {len(pdf_files)} PDF files in {pdf_folder}")
    
    # Create output folder
    os.makedirs(output_folder, exist_ok=True)
    
    # Process each PDF
    for filename in pdf_files:
        pdf_path = os.path.join(pdf_folder, filename)
        result = process_single_pdf(pdf_path, output_folder)
        results.append(result)
        
        if result.success:
            logger.info(f"  ✅ {filename}: {result.row_count} rows → {result.output_path}")
        else:
            logger.error(f"  ❌ {filename}: {result.error}")
    
    # Summary
    successful = sum(1 for r in results if r.success)
    failed = sum(1 for r in results if not r.success)
    logger.info(f"\n✅ Processed {successful} files successfully, {failed} failed")
    
    return results


def main() -> int:
    """Main entry point.
    
    Returns:
        Exit code (0 = success, 1 = failure)
    """
    try:
        # Setup logging
        Config.setup_logging()
        logger.info("Starting BCA statement extractor")
        
        # Validate configuration
        Config.validate_paths()
        
        # Process PDFs
        results = process_all_pdfs(Config.PDF_FOLDER, Config.OUTPUT_FOLDER)
        
        # Return status
        failed_count = sum(1 for r in results if not r.success)
        return 1 if failed_count > 0 else 0
        
    except Exception as e:
        logger.exception(f"Fatal error: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())