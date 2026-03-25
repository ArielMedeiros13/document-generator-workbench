# Document Generator Workbench

A browser-based document processing tool that imports PDFs, extracts content, and generates formatted documents with template support.

## Problem

Manual document creation and formatting is time-consuming, especially when working with existing PDF content that needs to be repurposed or enhanced with consistent styling.

## Solution

A client-side workbench that:
- Imports PDF documents using PDF.js
- Extracts and maps content intelligently
- Applies professional templates by document type
- Exports polished, print-ready PDFs

## Stack

- **Frontend**: Vanilla JavaScript (ES6+)
- **PDF Processing**: PDF.js 3.11.174 with CDN fallbacks
- **Styling**: Pure CSS with Nunito font family
- **Export**: Browser print API with CSS print optimization

## Highlights

- **Robust PDF Loading**: Dual-CDN fallback system ensures PDF.js always loads
- **Heuristic Mapping**: Smart content detection and field mapping
- **Template System**: Type-aware templates for consistent formatting
- **Print Optimization**: CSS-based print styling for professional output
- **Cache Management**: Version-controlled asset loading with cache-busting

## Features

- PDF import with drag-and-drop support
- Visual content mapping interface
- Real-time preview
- Multiple export formats
- Template customization
- Batch processing capabilities
