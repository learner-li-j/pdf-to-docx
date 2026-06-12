#!/usr/bin/env python
"""Setup configuration for PDF to DOCX converter"""

from setuptools import setup, find_packages
from pathlib import Path

# Read README
this_directory = Path(__file__).parent
long_description = (this_directory / "README.md").read_text(encoding="utf-8")

setup(
    name="pdf-to-docx-converter",
    version="1.0.0",
    author="learner-li-j",
    description="High-quality PDF to DOCX converter with multi-engine support",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/learner-li-j/pdf-to-docx",
    project_urls={
        "Bug Tracker": "https://github.com/learner-li-j/pdf-to-docx/issues",
        "Documentation": "https://github.com/learner-li-j/pdf-to-docx",
    },
    classifiers=[
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Development Status :: 4 - Beta",
        "Intended Audience :: Developers",
        "Topic :: Office/Business",
        "Topic :: Multimedia :: Graphics :: Viewers",
    ],
    python_requires=">=3.7",
    install_requires=[
        "python-docx>=0.8.11",
        "pdfplumber>=0.9.0",
        "pdf2image>=1.16.0",
        "pytesseract>=0.3.10",
        "pymupdf>=1.23.0",
        "Pillow>=9.0.0",
    ],
    extras_require={
        "dev": [
            "pytest>=7.0.0",
            "pytest-cov>=4.0.0",
            "black>=22.0.0",
            "flake8>=4.0.0",
            "mypy>=0.950",
        ],
        "gui": ["tkinter"],
    },
    entry_points={
        "console_scripts": [
            "pdf2docx=pdf_to_docx_converter:main",
            "pdf2docx-gui=gui_converter:main",
        ],
    },
    packages=find_packages(),
    include_package_data=True,
)