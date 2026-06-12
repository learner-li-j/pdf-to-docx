#!/usr/bin/env python
"""Basic usage examples for PDF to DOCX converter"""

from pdf_to_docx_converter import (
    convert_pdf_to_docx,
    PDFToDOCXConverter,
    PDFToDOCXError,
    get_pdf_converter_info,
)
from pathlib import Path


def example_1_simple_conversion():
    """
    Example 1: Simple conversion with automatic method selection
    """
    print("\n" + "="*50)
    print("Example 1: Simple Conversion")
    print("="*50)
    
    try:
        # Convert PDF to DOCX
        # Output will be saved in the same directory as input
        docx_path = convert_pdf_to_docx('sample.pdf')
        print(f"✅ Success: {docx_path}")
    except PDFToDOCXError as e:
        print(f"❌ Error: {e}")


def example_2_custom_output():
    """
    Example 2: Conversion with custom output path
    """
    print("\n" + "="*50)
    print("Example 2: Custom Output Path")
    print("="*50)
    
    try:
        docx_path = convert_pdf_to_docx(
            input_path='sample.pdf',
            output_path='output/converted.docx'
        )
        print(f"✅ Success: {docx_path}")
    except PDFToDOCXError as e:
        print(f"❌ Error: {e}")


def example_3_specific_method():
    """
    Example 3: Force specific conversion method
    """
    print("\n" + "="*50)
    print("Example 3: Specific Conversion Method")
    print("="*50)
    
    try:
        # Use pdfplumber method (fast, good for regular PDFs)
        docx_path = convert_pdf_to_docx(
            input_path='sample.pdf',
            method='pdfplumber'
        )
        print(f"✅ pdfplumber: {docx_path}")
    except PDFToDOCXError as e:
        print(f"❌ pdfplumber failed: {e}")
    
    try:
        # Use OCR method (for scanned PDFs)
        docx_path = convert_pdf_to_docx(
            input_path='scan.pdf',
            method='ocr'
        )
        print(f"✅ OCR: {docx_path}")
    except PDFToDOCXError as e:
        print(f"❌ OCR failed: {e}")
    
    try:
        # Use LibreOffice (best compatibility)
        docx_path = convert_pdf_to_docx(
            input_path='sample.pdf',
            method='libreoffice'
        )
        print(f"✅ LibreOffice: {docx_path}")
    except PDFToDOCXError as e:
        print(f"❌ LibreOffice failed: {e}")


def example_4_custom_timeout():
    """
    Example 4: Custom timeout for large files
    """
    print("\n" + "="*50)
    print("Example 4: Custom Timeout")
    print("="*50)
    
    try:
        # Increase timeout to 5 minutes for large PDFs
        docx_path = convert_pdf_to_docx(
            input_path='large_file.pdf',
            timeout=300  # 5 minutes
        )
        print(f"✅ Success: {docx_path}")
    except PDFToDOCXError as e:
        print(f"❌ Error: {e}")


def example_5_class_usage():
    """
    Example 5: Using PDFToDOCXConverter class directly
    """
    print("\n" + "="*50)
    print("Example 5: Class Usage")
    print("="*50)
    
    # Create converter instance
    converter = PDFToDOCXConverter(timeout=120, prefer_ocr=False)
    
    try:
        # Convert with auto method selection
        docx_path = converter.convert(
            input_path='sample.pdf',
            output_path='output/result.docx',
            method='auto'
        )
        print(f"✅ Conversion successful: {docx_path}")
    except PDFToDOCXError as e:
        print(f"❌ Conversion failed: {e}")
        # Print detailed error information
        if converter.conversion_errors:
            print("\nDetailed errors:")
            for error in converter.conversion_errors:
                print(f"  - {error}")


def example_6_batch_conversion():
    """
    Example 6: Batch conversion of multiple PDFs
    """
    print("\n" + "="*50)
    print("Example 6: Batch Conversion")
    print("="*50)
    
    # Find all PDFs in a directory
    pdf_folder = Path('pdfs')
    pdf_files = list(pdf_folder.glob('*.pdf'))
    
    print(f"Found {len(pdf_files)} PDF files")
    
    converter = PDFToDOCXConverter(timeout=120)
    output_folder = Path('output')
    output_folder.mkdir(exist_ok=True)
    
    successful = 0
    failed = 0
    
    for pdf_file in pdf_files:
        try:
            output_path = output_folder / (pdf_file.stem + '.docx')
            docx_path = converter.convert(
                input_path=str(pdf_file),
                output_path=str(output_path)
            )
            print(f"✅ {pdf_file.name} -> {output_path.name}")
            successful += 1
        except PDFToDOCXError as e:
            print(f"❌ {pdf_file.name}: {str(e)[:50]}...")
            failed += 1
    
    print(f"\n📊 Results: {successful} successful, {failed} failed")


def example_7_check_available_methods():
    """
    Example 7: Check available conversion methods
    """
    print("\n" + "="*50)
    print("Example 7: Check Available Methods")
    print("="*50)
    
    info = get_pdf_converter_info()
    
    print("\n📦 Available Conversion Methods:")
    for method, available in info.items():
        status = "✅" if available else "❌"
        print(f"  {status} {method}")
    
    # Use converter to get detailed info
    converter = PDFToDOCXConverter()
    available_methods = converter._get_available_methods()
    
    print(f"\n🔄 Conversion will try: {', '.join(available_methods) if available_methods else 'None'}")


def example_8_error_handling():
    """
    Example 8: Comprehensive error handling
    """
    print("\n" + "="*50)
    print("Example 8: Error Handling")
    print("="*50)
    
    try:
        # Try to convert non-existent file
        docx_path = convert_pdf_to_docx('/nonexistent/file.pdf')
    except PDFToDOCXError as e:
        print(f"❌ PDFToDOCXError caught:")
        print(f"   {e}")
    except Exception as e:
        print(f"❌ Unexpected error:")
        print(f"   {type(e).__name__}: {e}")


if __name__ == '__main__':
    print("\n" + "#"*50)
    print("# PDF to DOCX Converter - Usage Examples")
    print("#"*50)
    
    # Run all examples
    example_1_simple_conversion()
    example_2_custom_output()
    example_3_specific_method()
    example_4_custom_timeout()
    example_5_class_usage()
    example_6_batch_conversion()
    example_7_check_available_methods()
    example_8_error_handling()
    
    print("\n" + "#"*50)
    print("# Examples Complete")
    print("#"*50)
