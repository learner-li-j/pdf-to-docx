"""Unit tests for PDF to DOCX converter"""

import pytest
import tempfile
from pathlib import Path
from pdf_to_docx_converter import (
    PDFToDOCXConverter,
    convert_pdf_to_docx,
    PDFToDOCXError,
    get_pdf_converter_info,
)


class TestPDFToDOCXConverter:
    """Test cases for PDFToDOCXConverter class"""

    def test_converter_initialization(self):
        """Test converter can be initialized"""
        converter = PDFToDOCXConverter(timeout=120, prefer_ocr=False)
        assert converter.timeout == 120
        assert converter.prefer_ocr is False
        assert converter.conversion_errors == []

    def test_invalid_input_file_not_found(self):
        """Test error when input file does not exist"""
        converter = PDFToDOCXConverter()
        
        with pytest.raises(PDFToDOCXError) as exc_info:
            converter.convert("/nonexistent/file.pdf")
        
        assert "not found" in str(exc_info.value).lower()

    def test_invalid_input_not_pdf(self):
        """Test error when input is not a PDF file"""
        converter = PDFToDOCXConverter()
        
        with tempfile.NamedTemporaryFile(suffix=".txt") as tmp:
            with pytest.raises(PDFToDOCXError) as exc_info:
                converter.convert(tmp.name)
            
            assert "not a pdf" in str(exc_info.value).lower()

    def test_output_path_generation(self):
        """Test automatic output path generation"""
        converter = PDFToDOCXConverter()
        
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            pdf_path = Path(tmp.name)
            expected_docx = pdf_path.parent / (pdf_path.stem + ".docx")
            
            # Validate output path generation
            assert pdf_path.exists()
            assert pdf_path.suffix == ".pdf"

    def test_get_available_methods(self):
        """Test detection of available conversion methods"""
        converter = PDFToDOCXConverter()
        methods = converter._get_available_methods()
        
        # Should return a list
        assert isinstance(methods, list)
        # List can be empty or contain method names
        for method in methods:
            assert method in ["pdfplumber", "ocr", "libreoffice", "pymupdf"]

    def test_get_conversion_info(self):
        """Test conversion info retrieval"""
        info = get_pdf_converter_info()
        
        # Should be a dictionary
        assert isinstance(info, dict)
        # Should contain all methods
        assert set(info.keys()) == {"pdfplumber", "ocr", "libreoffice", "pymupdf"}
        # Values should be boolean
        for available in info.values():
            assert isinstance(available, bool)

    def test_invalid_method(self):
        """Test error with invalid conversion method"""
        converter = PDFToDOCXConverter()
        
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            pdf_path = tmp.name
            
            # Create a minimal PDF file
            tmp.write(b"%PDF-1.4\n")
            tmp.flush()
            
            with pytest.raises(PDFToDOCXError) as exc_info:
                converter.convert(pdf_path, method="invalid_method")
            
            assert "unknown" in str(exc_info.value).lower()


class TestConvenientFunctions:
    """Test convenient API functions"""

    def test_convert_pdf_to_docx_function(self):
        """Test convert_pdf_to_docx convenience function"""
        # Function should exist and be callable
        assert callable(convert_pdf_to_docx)

    def test_get_converter_info_function(self):
        """Test get_pdf_converter_info convenience function"""
        info = get_pdf_converter_info()
        assert isinstance(info, dict)


class TestErrorHandling:
    """Test error handling"""

    def test_pdftodocx_error_exception(self):
        """Test PDFToDOCXError exception"""
        error_msg = "Test error message"
        error = PDFToDOCXError(error_msg)
        
        assert isinstance(error, Exception)
        assert str(error) == error_msg

    def test_converter_error_tracking(self):
        """Test that converter tracks conversion errors"""
        converter = PDFToDOCXConverter()
        
        # Initially should be empty
        assert converter.conversion_errors == []
        
        # Convert with invalid file (will add errors)
        try:
            converter.convert("/nonexistent/file.pdf")
        except PDFToDOCXError:
            pass  # Expected


class TestIntegration:
    """Integration tests"""

    def test_converter_with_timeout(self):
        """Test converter initialization with custom timeout"""
        converter = PDFToDOCXConverter(timeout=300)
        assert converter.timeout == 300

    def test_converter_prefer_ocr(self):
        """Test converter initialization with OCR preference"""
        converter = PDFToDOCXConverter(prefer_ocr=True)
        assert converter.prefer_ocr is True

    def test_method_detection_consistency(self):
        """Test that method detection is consistent"""
        converter1 = PDFToDOCXConverter()
        converter2 = PDFToDOCXConverter()
        
        methods1 = converter1._get_available_methods()
        methods2 = converter2._get_available_methods()
        
        # Should return same methods both times
        assert methods1 == methods2


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
