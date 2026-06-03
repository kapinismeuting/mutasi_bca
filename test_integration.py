#!/usr/bin/env python3
"""
Integration tests for the BCA bank statement extractor.
"""

import pytest
import os
import tempfile
from pathlib import Path
from mutasi import process_single_pdf, process_all_pdfs
from mutasi_by_year import process_all_pdfs_to_single_excel


class TestIntegration:
    """Integration tests for the complete workflow."""
    
    def test_process_all_pdfs_empty_folder(self):
        """Test processing empty PDF folder."""
        with tempfile.TemporaryDirectory() as tmpdir:
            output_dir = os.path.join(tmpdir, "output")
            
            # Should not raise error, but return empty results
            results = process_all_pdfs(tmpdir, output_dir)
            assert len(results) == 0
    
    def test_process_all_pdfs_missing_folder(self):
        """Test processing non-existent PDF folder."""
        with pytest.raises(ValueError):
            process_all_pdfs("/nonexistent/folder", "/tmp/output")
    
    def test_consolidated_processing_empty_folder(self):
        """Test consolidated processing of empty folder."""
        with tempfile.TemporaryDirectory() as tmpdir:
            output_dir = os.path.join(tmpdir, "output")
            
            # Should not raise error
            results = process_all_pdfs_to_single_excel(tmpdir, output_dir)
            assert len(results) == 0


class TestErrorHandling:
    """Test error handling in processing."""
    
    def test_process_missing_pdf(self):
        """Test processing non-existent PDF file."""
        result = process_single_pdf("/nonexistent/file.pdf", "/tmp/output")
        assert result.success is False
        assert "not found" in result.error.lower()
    
    def test_process_all_pdfs_invalid_pdf(self):
        """Test processing folder with invalid PDF."""
        with tempfile.TemporaryDirectory() as tmpdir:
            output_dir = os.path.join(tmpdir, "output")
            
            # Create a non-PDF file
            invalid_file = os.path.join(tmpdir, "notapdf.pdf")
            with open(invalid_file, 'w') as f:
                f.write("This is not a PDF")
            
            # Should handle gracefully
            results = process_all_pdfs(tmpdir, output_dir)
            # Will have one failed result
            assert len(results) >= 1


class TestOutputCreation:
    """Test output file creation."""
    
    def test_output_folder_created(self):
        """Test that output folder is created if missing."""
        with tempfile.TemporaryDirectory() as tmpdir:
            output_dir = os.path.join(tmpdir, "nested/output/folder")
            
            # Create empty folder for PDFs
            pdf_dir = os.path.join(tmpdir, "pdfs")
            os.makedirs(pdf_dir)
            
            # Process (will create output folder)
            process_all_pdfs(pdf_dir, output_dir)
            
            # Folder should exist (even if empty)
            assert os.path.exists(output_dir)


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
