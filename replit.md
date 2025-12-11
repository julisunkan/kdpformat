# KDP Book Formatter

A professional-grade Progressive Web App (PWA) for formatting manuscripts according to Amazon Kindle Direct Publishing (KDP) standards.

## Overview

This Flask-based PWA helps authors prepare their manuscripts for publishing on Amazon KDP by:
- Applying proper formatting (margins, fonts, spacing)
- Generating front matter (title page, copyright, dedication)
- Creating a Word-compatible table of contents
- Checking image DPI for print quality
- Converting to PDF for print books

## Project Structure

```
├── app.py                 # Main Flask application
├── services/
│   ├── __init__.py
│   ├── formatter.py       # Core KDP formatting logic
│   ├── frontmatter.py     # Title/copyright/dedication pages
│   ├── dpi_checker.py     # Image DPI validation
│   └── pdf_exporter.py    # DOCX to PDF conversion
├── templates/
│   ├── upload.html        # File upload interface
│   ├── results.html       # Processing results page
│   └── offline.html       # PWA offline page
├── static/
│   ├── style.css          # Dark mode styling
│   ├── script.js          # Drag-and-drop functionality
│   ├── manifest.json      # PWA manifest
│   ├── sw.js              # Service worker
│   └── icons/             # PWA icons (72-512px, standard & maskable)
└── uploads/               # Temporary file storage
```

## PWA Features

- **Installable**: Can be installed on desktop and mobile devices
- **Offline support**: Shows offline page when no internet connection
- **App-like experience**: Standalone display mode without browser chrome
- **Icons**: Multiple sizes including maskable icons for Android

## Features

### Formatting Options
- **Trim Sizes**: 6×9, 5×8, 8.5×11 inches
- **Margins**: Top/Bottom 1", Inside 0.85", Outside 0.6"
- **Print Mode**: Mirrored margins for physical books
- **Line Spacing**: 1.15, 1.25, or 1.5

### Text Formatting
- Font: Georgia 11pt
- First-line indent: 0.25"
- Paragraph spacing after: 6pt
- Automatic cleanup of extra spaces, tabs, line breaks

### Chapter Detection
- Heading 1 paragraphs treated as chapter titles
- Page break before each chapter
- Center-aligned chapter headings

### Front Matter
- Title page with book title and author name
- Copyright page with standard notice
- Dedication page

### Table of Contents
- Word-compatible dynamic TOC using field codes
- Renders when opened in Microsoft Word

### Image DPI Checking
- Scans all embedded images
- Warns if DPI < 300 (print minimum)

### PDF Export
- Uses LibreOffice headless for conversion
- Output: PRINT_<uuid>.pdf

## Running the Application

The application runs on port 5000:
```bash
python app.py
```

## Dependencies

- Flask - Web framework
- python-docx - DOCX manipulation
- Pillow - Image DPI checking
- LibreOffice - PDF conversion (system dependency)

## User Preferences

- Dark mode UI by default
- Drag-and-drop file upload
- Immediate download after processing
