# PDF to PowerPoint Converter

This Python script converts PDF files to PowerPoint presentations with automatic orientation handling.

## Features
- Converts PDF pages to PowerPoint slides
- Automatically detects page orientation
  - Vertical pages (height > width) are split into two slides
  - Horizontal/square pages are kept as single slides
- Preserves aspect ratio and centers images
- Shows real-time conversion progress
- Handles LaTeX-generated PDFs

## Requirements
- Python 3.x
- pdf2image
- python-pptx
- poppler (for pdf2image)

## Installation
```bash
pip install pdf2image python-pptx
```

## Usage
```bash
python convert.py input.pdf output.pptx
```

## Example
```bash
python convert.py presentation.pdf slides.pptx
```

## Notes
- Input file must be a PDF
- Output file must be a PPTX
- Temporary PNG files will be created and deleted during conversion
- Vertical pages will create two slides per page
- Horizontal/square pages will create one slide per page
