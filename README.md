# PDF to PowerPoint Converter

This Python script converts PDF files to PowerPoint presentations.

## Features
- Converts PDF pages to PowerPoint slides
- Preserves text formatting
- Handles LaTeX-generated PDFs

## Requirements
- Python 3.x
- pdf2image
- python-pptx

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
