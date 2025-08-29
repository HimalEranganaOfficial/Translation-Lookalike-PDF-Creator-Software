"""
Comprehensive Document Converter

This script performs a multi-step conversion process:
1. Converts a .docx file to a .pdf file.
2. Converts the created .pdf file into a series of .jpg images,
   saving them to a new subfolder.
3. Renames the .jpg files sequentially (1.jpg, 2.jpg, etc.).
4. Converts the .jpg images back into a new "scanned" .pdf file.
"""

import os
from pathlib import Path
from docx2pdf import convert
from pdf2image import convert_from_path
import img2pdf

# --- 1. SET YOUR FILE PATH ---
# Place the pathway to your .docx file here.
# Make sure to use a raw string (r"...") or double backslashes (\\)
# for Windows paths to avoid issues with escape characters.
input_docx_path = r"C:\Users\user\Downloads\6\2222.docx"

def convert_docx_to_pdf(docx_path):
    """Converts a .docx file to a .pdf file."""
    pdf_path = Path(docx_path).with_suffix('.pdf')
    try:
        print(f"Starting conversion of '{docx_path}' to PDF...")
        convert(str(docx_path), str(pdf_path))
        print(f"DOCX to PDF conversion successful! PDF saved to '{pdf_path}'")
        return pdf_path
    except Exception as e:
        print(f"An error occurred during DOCX to PDF conversion: {e}")
        return None

def convert_pdf_to_images(pdf_path):
    """Converts a PDF file to a series of JPG images."""
    parent_dir = pdf_path.parent
    base_name = pdf_path.stem
    image_folder = parent_dir / f"{base_name}_images"
    os.makedirs(image_folder, exist_ok=True)
    
    print(f"Converting '{pdf_path}' to JPGs and saving to '{image_folder}'...")
    try:
        images = convert_from_path(str(pdf_path))
        image_paths = []
        for i, image in enumerate(images):
            image_name = f"{i + 1}.jpg"
            image_path = image_folder / image_name
            image.save(image_path, "JPEG")
            image_paths.append(image_path)
        print(f"PDF to JPG conversion successful! Saved {len(image_paths)} images.")
        return image_paths
    except Exception as e:
        print(f"An error occurred during PDF to JPG conversion: {e}")
        print("Please ensure you have Poppler installed and in your system PATH.")
        return []

def create_pdf_from_images(image_paths, original_pdf_path):
    """Creates a new PDF from a list of JPG images."""
    if not image_paths:
        print("No images found to create a new PDF.")
        return None

    # Sort images numerically to ensure correct page order
    image_paths.sort(key=lambda p: int(p.stem))
    
    # Create the new PDF file name
    parent_dir = original_pdf_path.parent
    base_name = original_pdf_path.stem
    new_pdf_path = parent_dir / f"scan_{base_name}.pdf"
    
    print(f"Creating new PDF from images and saving to '{new_pdf_path}'...")
    try:
        with open(new_pdf_path, "wb") as f:
            f.write(img2pdf.convert([str(p) for p in image_paths]))
        print(f"New PDF created successfully! Saved to '{new_pdf_path}'")
        return new_pdf_path
    except Exception as e:
        print(f"An error occurred during image to PDF conversion: {e}")
        return None

# --- Main script execution ---
if __name__ == "__main__":
    
    # Step 1: Convert DOCX to PDF
    pdf_file = convert_docx_to_pdf(input_docx_path)

    if pdf_file:
        # Step 2: Convert PDF to JPGs
        jpg_files = convert_pdf_to_images(pdf_file)

        # Step 3: Convert JPGs back to a new PDF
        if jpg_files:
            create_pdf_from_images(jpg_files, pdf_file)
