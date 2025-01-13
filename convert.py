import os
from pdf2image import convert_from_path
from pptx import Presentation
from pptx.util import Inches

def pdf_to_ppt(pdf_path, ppt_path, dpi=300):
    """
    Converts a PDF (created from LaTeX) into a PowerPoint presentation with each page as an image.

    Args:
        pdf_path (str): Path to the input PDF file.
        ppt_path (str): Path to save the output PPT file.
        dpi (int): DPI for converting PDF pages to images.
    """
    # Create a PowerPoint presentation
    presentation = Presentation()

    # Convert PDF pages to images
    images = convert_from_path(pdf_path, dpi=dpi)

    for idx, img in enumerate(images):
        # Save each image temporarily
        img_path = f"temp_page_{idx + 1}.png"
        img.save(img_path, "PNG")

        # Add a slide for each image
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])  # Blank slide layout
        slide.shapes.add_picture(img_path, Inches(0), Inches(0), width=Inches(10), height=Inches(7.5))

        # Remove the temporary image
        os.remove(img_path)

    # Save the presentation
    presentation.save(ppt_path)
    print(f"PowerPoint saved at {ppt_path}")

# Example usage
pdf_path = "AI_02_Fundamentals_of_AI.pdf"
ppt_path = "output_presentation.pptx"
pdf_to_ppt(pdf_path, ppt_path)