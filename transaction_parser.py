#!/usr/bin/env python3
"""
BCA bank statement transaction parser.
Extracts transactions from PDF text with robust error handling.
"""

import re
import pdfplumber
from typing import List, Dict, Tuple, Optional
from datetime import datetime
from config import Config, get_logger

logger = get_logger('mutasi.parser')


class Transaction:
    """Type-safe transaction representation."""
    
    def __init__(
        self,
        day: int,
        month: int,
        description: str,
        debit: str = "",
        credit: str = "",
        balance: str = ""
    ):
        """Initialize a transaction.
        
        Args:
            day: Day of month (1-31)
            month: Month (1-12)
            description: Transaction description
            debit: Debit amount or empty string
            credit: Credit amount or empty string
            balance: Balance amount or empty string
        """
        self.day = day
        self.month = month
        self.description = description
        self.debit = debit
        self.credit = credit
        self.balance = balance
    
    def to_dict(self) -> Dict[str, any]:
        """Convert to dictionary representation."""
        return {
            'tanggal': self.day,
            'bulan': self.month,
            'keterangan': self.description,
            'db': self.debit,
            'cr': self.credit,
            'saldo': self.balance
        }
    
    def validate(self) -> Tuple[bool, Optional[str]]:
        """Validate transaction data.
        
        Returns:
            Tuple of (is_valid, error_message)
        """
        if not (1 <= self.day <= 31):
            return False, f"Invalid day: {self.day}"
        
        if not (1 <= self.month <= 12):
            return False, f"Invalid month: {self.month}"
        
        if not isinstance(self.description, str):
            return False, "Description must be string"
        
        return True, None


# Pre-compiled regex patterns (avoid recompilation)
DATE_PATTERN = re.compile(Config.DATE_PATTERN)
DB_PATTERN = re.compile(Config.DB_PATTERN)
CR_PATTERN = re.compile(Config.CR_PATTERN)
SALDO_PATTERN = re.compile(Config.SALDO_PATTERN)
SUMMARY_PATTERN = re.compile(Config.SUMMARY_PATTERN)


def extract_date(line: str) -> Optional[Tuple[int, int]]:
    """Extract day and month from line.
    
    Args:
        line: Text line to parse
        
    Returns:
        Tuple of (day, month) or None if not found
    """
    match = DATE_PATTERN.match(line.strip())
    if match:
        day = int(match.group(1))
        month = int(match.group(2))
        return day, month
    return None


def extract_amounts(text: str, first_line: str) -> Tuple[str, str]:
    """Extract debit and credit amounts.
    
    Args:
        text: Full transaction text
        first_line: First line of transaction
        
    Returns:
        Tuple of (debit, credit) amounts
    """
    debit = ""
    credit = ""
    
    # Try to find debit amount
    db_match = DB_PATTERN.search(text)
    if db_match:
        debit = db_match.group(1)
    else:
        # If no debit, try credit from first line
        cr_matches = CR_PATTERN.findall(first_line)
        if cr_matches:
            credit = cr_matches[0]
    
    return debit, credit


def extract_balance(first_line: str) -> str:
    """Extract balance amount.
    
    Args:
        first_line: First line of transaction
        
    Returns:
        Balance amount or empty string
    """
    match = SALDO_PATTERN.search(first_line)
    if match:
        return match.group(1)
    return ""


def parse_transaction_block(
    lines: List[str],
    start_index: int
) -> Tuple[Optional[Transaction], int]:
    """Parse a transaction block from lines.
    
    Args:
        lines: List of all text lines
        start_index: Index to start parsing from
        
    Returns:
        Tuple of (transaction, next_index) or (None, next_index)
    """
    if start_index >= len(lines):
        return None, start_index
    
    line = lines[start_index].strip()
    if not line:
        return None, start_index + 1
    
    # Try to extract date
    date_info = extract_date(line)
    if not date_info:
        return None, start_index
    
    day, month = date_info
    first_line = line
    
    # Collect description lines
    date_str = DATE_PATTERN.match(line).group(0)
    desc_lines = [line[len(date_str):].strip()]
    
    j = start_index + 1
    while j < len(lines):
        next_line = lines[j].strip()
        
        if not next_line:
            j += 1
            continue
        
        # Stop at next transaction or summary
        if DATE_PATTERN.match(next_line) or SUMMARY_PATTERN.match(next_line):
            break
        
        desc_lines.append(next_line)
        j += 1
    
    full_desc = "\n".join(desc_lines).strip()
    
    # Extract amounts
    debit, credit = extract_amounts(full_desc, first_line)
    balance = extract_balance(first_line)
    
    # Create transaction
    transaction = Transaction(
        day=day,
        month=month,
        description=full_desc,
        debit=debit,
        credit=credit,
        balance=balance
    )
    
    # Validate
    is_valid, error = transaction.validate()
    if not is_valid:
        logger.warning(f"Invalid transaction at line {start_index}: {error}")
    
    return transaction, j


def parse_summary_line(line: str) -> Optional[Dict]:
    """Parse summary line.
    
    Args:
        line: Summary line to parse
        
    Returns:
        Dictionary with summary data or None
    """
    match = SUMMARY_PATTERN.match(line.strip())
    if match:
        amount_match = re.search(r'(\d{1,3}(?:,\d{3})*\.\d{2})', line)
        amount = amount_match.group(1) if amount_match else ""
        
        return {
            'tanggal': "",
            'bulan': "",
            'keterangan': line.strip(),
            'db': "",
            'cr': "",
            'saldo': amount
        }
    return None


def parse_bca_transactions(pdf_path: str) -> List[Dict]:
    """Extract transactions from BCA PDF statement.
    
    Args:
        pdf_path: Path to PDF file to parse
        
    Returns:
        List of transaction dictionaries
        
    Raises:
        FileNotFoundError: If PDF file does not exist
        ValueError: If file is too large
        pdfplumber.PDFException: If PDF cannot be read
    """
    import os
    
    # Validate file exists
    if not os.path.exists(pdf_path):
        logger.error(f"PDF file not found: {pdf_path}")
        raise FileNotFoundError(f"PDF not found: {pdf_path}")
    
    # Check file size
    file_size = os.path.getsize(pdf_path)
    if file_size > Config.MAX_PDF_SIZE:
        logger.error(f"PDF file too large ({file_size} bytes): {pdf_path}")
        raise ValueError(f"PDF file exceeds maximum size ({Config.MAX_PDF_SIZE} bytes)")
    
    transactions = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            logger.info(f"Parsing PDF with {len(pdf.pages)} pages: {pdf_path}")
            
            for page_num, page in enumerate(pdf.pages, 1):
                try:
                    text = page.extract_text()
                    if not text:
                        logger.warning(f"Page {page_num} is empty or unreadable")
                        continue
                    
                    lines = text.split('\n')
                    i = 0
                    
                    while i < len(lines):
                        line = lines[i].strip()
                        
                        if not line:
                            i += 1
                            continue
                        
                        # Try to parse as transaction
                        transaction, next_i = parse_transaction_block(lines, i)
                        if transaction:
                            transactions.append(transaction.to_dict())
                            i = next_i
                            continue
                        
                        # Try to parse as summary
                        summary = parse_summary_line(line)
                        if summary:
                            transactions.append(summary)
                            i += 1
                            continue
                        
                        i += 1
                
                except Exception as e:
                    logger.error(f"Error parsing page {page_num}: {e}")
                    raise
    
    except pdfplumber.PDFException as e:
        logger.error(f"Invalid PDF file: {pdf_path} - {e}")
        raise
    except Exception as e:
        logger.error(f"Error parsing PDF: {pdf_path} - {e}")
        raise
    
    logger.info(f"Successfully parsed {len(transactions)} transactions from {pdf_path}")
    return transactions
