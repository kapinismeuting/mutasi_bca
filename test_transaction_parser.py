#!/usr/bin/env python3
"""
Unit tests for transaction parser module.
"""

import pytest
from transaction_parser import (
    Transaction, extract_date, extract_amounts, extract_balance,
    parse_summary_line
)


class TestTransaction:
    """Test Transaction class."""
    
    def test_transaction_valid(self):
        """Test creating a valid transaction."""
        tx = Transaction(
            day=15,
            month=1,
            description="Transfer",
            debit="1,000.00"
        )
        assert tx.day == 15
        assert tx.month == 1
        assert tx.description == "Transfer"
        assert tx.debit == "1,000.00"
    
    def test_transaction_to_dict(self):
        """Test converting transaction to dictionary."""
        tx = Transaction(
            day=15,
            month=1,
            description="Transfer",
            debit="1,000.00",
            credit="",
            balance="5,000.00"
        )
        data = tx.to_dict()
        assert data['tanggal'] == 15
        assert data['bulan'] == 1
        assert data['keterangan'] == "Transfer"
        assert data['db'] == "1,000.00"
        assert data['saldo'] == "5,000.00"
    
    def test_transaction_validate_valid(self):
        """Test validation of valid transaction."""
        tx = Transaction(day=15, month=6, description="Test")
        is_valid, error = tx.validate()
        assert is_valid is True
        assert error is None
    
    def test_transaction_validate_invalid_day(self):
        """Test validation of invalid day."""
        tx = Transaction(day=32, month=1, description="Test")
        is_valid, error = tx.validate()
        assert is_valid is False
        assert error is not None
    
    def test_transaction_validate_invalid_month(self):
        """Test validation of invalid month."""
        tx = Transaction(day=15, month=13, description="Test")
        is_valid, error = tx.validate()
        assert is_valid is False
        assert error is not None


class TestDateExtraction:
    """Test date extraction function."""
    
    def test_extract_date_valid(self):
        """Test extracting valid date."""
        result = extract_date("15/06 Transfer 1,000.00")
        assert result == (15, 6)
    
    def test_extract_date_single_digit(self):
        """Test extracting date with single digit day."""
        result = extract_date("05/01 Transfer")
        assert result == (5, 1)
    
    def test_extract_date_invalid(self):
        """Test extracting invalid date."""
        result = extract_date("Transfer 1,000.00")
        assert result is None
    
    def test_extract_date_no_leading_zeros(self):
        """Test that date must start at beginning of line."""
        result = extract_date(" 15/06 Transfer")
        assert result is None  # Should not match after leading space


class TestAmountExtraction:
    """Test amount extraction functions."""
    
    def test_extract_balance_valid(self):
        """Test extracting balance amount."""
        result = extract_balance("15/06 Transfer 1,234.56")
        assert result == "1,234.56"
    
    def test_extract_balance_missing(self):
        """Test extracting missing balance."""
        result = extract_balance("15/06 Transfer")
        assert result == ""
    
    def test_extract_amounts_with_debit(self):
        """Test extracting amounts with debit."""
        db, cr = extract_amounts("500.00 DB", "15/06 Transfer 1,234.56")
        assert db == "500.00"
        assert cr == ""
    
    def test_extract_amounts_with_credit(self):
        """Test extracting amounts with credit."""
        db, cr = extract_amounts("Transfer 1,000.00", "15/06 Transfer 1,000.00")
        assert db == ""
        assert cr == "1,000.00"


class TestSummaryParsing:
    """Test summary line parsing."""
    
    def test_parse_summary_saldo_awal(self):
        """Test parsing SALDO AWAL summary."""
        result = parse_summary_line("SALDO AWAL: 1,000,000.00")
        assert result is not None
        assert result['keterangan'] == "SALDO AWAL: 1,000,000.00"
        assert result['saldo'] == "1,000,000.00"
    
    def test_parse_summary_mutasi_cr(self):
        """Test parsing MUTASI CR summary."""
        result = parse_summary_line("MUTASI CR: 500,000.00")
        assert result is not None
        assert result['saldo'] == "500,000.00"
    
    def test_parse_summary_invalid(self):
        """Test parsing invalid summary."""
        result = parse_summary_line("Regular transaction 1,000.00")
        assert result is None


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
