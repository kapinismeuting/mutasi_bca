#!/usr/bin/env python3
"""
Extract BCA bank statement transactions to consolidated Excel.
Combines all PDFs into a single Excel file with separate sheets for each month.

Configuration via environment variables:
- PDF_FOLDER: Input directory with PDFs
- OUTPUT_FOLDER: Output directory for Excel files
- LOG_LEVEL: Logging level (DEBUG, INFO, WARNING, ERROR)
- LOG_FILE: Optional log file path
"""

import os
import sys
from pathlib import Path
from typing import Dict, List, Optional
from dataclasses import dataclass

from config import Config, get_logger
from transaction_parser import parse_bca_transactions
from excel_writer import write_multiple_sheets_to_excel

logger = get_logger('mutasi.yearly')


@dataclass
class ProcessResult:
    """Result of processing a single PDF."""
    success: bool
    filename: str
    sheet_name: str
    row_count: int = 0
    error: Optional[str] = None


def process_pdf_for_sheet(pdf_path: str) -> Optional[ProcessResult]:
    """Process a single PDF file and extract sheet name from filename.
    
    Args:
        pdf_path: Path to PDF file
        
    Returns:
        ProcessResult or None if processing fails
    """
    filename = os.path.basename(pdf_path)
    sheet_name = os.path.splitext(filename)[0][:Config.SHEET_NAME_LENGTH].upper()
    
    try:
        logger.info(f"Processing: {filename}")
        
        # Parse transactions
        transactions = parse_bca_transactions(pdf_path)
        
        return ProcessResult(
            success=True,
            filename=filename,
            sheet_name=sheet_name,
            row_count=len(transactions)
        )
        
    except FileNotFoundError as e:
        error_msg = f"PDF file not found: {e}"
        logger.error(error_msg)
        return ProcessResult(
            success=False,
            filename=filename,
            sheet_name=sheet_name,
            error=error_msg
        )
    
    except ValueError as e:
        error_msg = f"Invalid input: {e}"
        logger.error(error_msg)
        return ProcessResult(
            success=False,
            filename=filename,
            sheet_name=sheet_name,
            error=error_msg
        )
    
    except Exception as e:
        error_msg = f"Error: {e}"
        logger.error(error_msg)
        return ProcessResult(
            success=False,
            filename=filename,
            sheet_name=sheet_name,
            error=error_msg
        )


def process_all_pdfs_to_single_excel(pdf_folder: str, output_folder: str) -> List[ProcessResult]:
    """Process all PDFs in folder and write to consolidated Excel file.
    
    Args:
        pdf_folder: Input directory with PDFs
        output_folder: Output directory for Excel file
        
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
    
    # Process each PDF and collect transactions
    transactions_by_sheet: Dict[str, List[Dict]] = {}
    
    for filename in pdf_files:
        pdf_path = os.path.join(pdf_folder, filename)
        result = process_pdf_for_sheet(pdf_path)
        
        if result is None:
            continue
        
        results.append(result)
        
        if result.success:
            # Get transactions for this PDF
            try:
                transactions = parse_bca_transactions(pdf_path)
                transactions_by_sheet[result.sheet_name] = transactions
                logger.info(f"  ✅ {filename}: {result.row_count} rows → Sheet '{result.sheet_name}'")
            except Exception as e:
                logger.error(f"  ❌ Failed to parse {filename}: {e}")
        else:
            logger.error(f"  ❌ {filename}: {result.error}")
    
    # Write consolidated Excel file
    if transactions_by_sheet:
        output_filename = os.path.basename(pdf_folder) + ".xlsx"
        output_path = os.path.join(output_folder, output_filename)
        
        try:
            write_multiple_sheets_to_excel(transactions_by_sheet, output_path)
            logger.info(f"✅ Consolidated Excel file created: {output_path}")
        except Exception as e:
            logger.error(f"Failed to write consolidated Excel: {e}")
    else:
        logger.warning("No transactions to write")
    
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
        logger.info("Starting BCA consolidated statement extractor")
        
        # Validate configuration
        Config.validate_paths()
        
        # Process PDFs
        results = process_all_pdfs_to_single_excel(Config.PDF_FOLDER, Config.OUTPUT_FOLDER)
        
        # Return status
        failed_count = sum(1 for r in results if not r.success)
        return 1 if failed_count > 0 else 0
        
    except Exception as e:
        logger.exception(f"Fatal error: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())