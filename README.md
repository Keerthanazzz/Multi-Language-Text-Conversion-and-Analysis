# Multi-Language-Text-Conversion-and-Analysis
# Multi-Language Text Conversion and Analysis

This project is a Streamlit-based application designed to extract, clean, structure, and translate text from PDF files. It integrates OCR capabilities using Tesseract, generative AI for text processing, and provides multi-language support with accuracy evaluation and PDF generation features.

## Features

- **PDF Text Extraction**: Extract text from PDF files using Tesseract OCR with support for multiple languages.
- **Text Cleaning and Structuring**: Clean and format extracted text into a structured format using Google Generative AI.
- **Language Detection**: Automatically detect the language of the extracted text.
- **Multi-Language Translation**: Translate text into various Indian languages.
- **Accuracy Evaluation**: Evaluate translation accuracy by back-translating and comparing with the original text.
- **PDF Generation**: Save translated text as a PDF document for download.

## Technologies Used

- **Streamlit**: For building the user interface.
- **Google Generative AI**: For text cleaning, structuring, and translation.
- **Tesseract OCR**: For extracting text from images in PDF files.
- **Poppler**: For converting PDF pages to images.
- **Python Libraries**: `pytesseract`, `pdf2image`, `langdetect`, `docx`, `docx2pdf`, `pythoncom`, and `win32com.client`.

## Setup Instructions

### Prerequisites

- Python 3.8 or later
- Tesseract OCR installed ([Download here](https://github.com/tesseract-ocr/tesseract))
- Poppler installed ([Download here](http://blog.alivate.com.au/poppler-windows/))
- Google Generative AI API key
- Microsoft Word (for DOCX to PDF conversion)

