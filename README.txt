Document Conversion Project

Overview
This project focuses on converting unstructured documents into rule-based structured documents. It supports multiple input formats including PDF, DOCX, DOC, and HTML files, using specialized tools for each format to ensure optimal conversion quality.

Tools Used
ABBYY FineReader: For PDF document conversion and OCR
COM Conversion: For handling Microsoft Word documents (.doc and .docx files)
Python: Core programming language for the project
html2docx: Python library for HTML to DOCX conversion

Features
Converts PDFs to structured documents using ABBYY's OCR capabilities
Processes Microsoft Word documents (.doc, .docx) using COM interfaces
Converts HTML documents to structured DOCX format
Applies rule-based structuring to standardize document formats

Requirements
Python 3.7+
ABBYY FineReader SDK
Microsoft Office (for COM conversion)