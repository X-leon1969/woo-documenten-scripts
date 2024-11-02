import sys
import os
from pdf2image import convert_from_path
from PIL import Image
import pytesseract
from PyPDF2 import PdfWriter, PdfReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader
import io
from pathlib import Path
import shutil

def setup_fonts():
    pdfmetrics.registerFont(TTFont('DejaVuSans', 'DejaVuSans.ttf'))

def is_pdf_searchable(pdf_path):
    try:
        reader = PdfReader(open(pdf_path, "rb"))
        if len(reader.pages) > 0:
            page = reader.pages[0]
            if page.extract_text().strip():
                return True
    except PyPDF2.errors.PdfReadError:
        print(f"Error reading {os.path.basename(pdf_path)}. Assuming it's not searchable.")
    except Exception as e:
        print(f"An unexpected error occurred while checking {os.path.basename(pdf_path)}: {e}")
    return False

def process_single_pdf(pdf_path, target_dir):
    if is_pdf_searchable(pdf_path):
        print(f"{os.path.basename(pdf_path)} is already searchable and selectable. Skipping.")
        return

    output_path = os.path.join(target_dir, os.path.basename(pdf_path))  # Save with original filename
    setup_fonts()

    # Use a lower DPI for conversion, e.g., 150 instead of 300
    images = convert_from_path(pdf_path, dpi=150)
    merger = PdfWriter()

    for idx, img in enumerate(images):
        print(f"Processing page {idx + 1} of {len(images)}")
        
        # OCR with lower resolution image
        data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT)
        h, w = img.height, img.width

        # Convert to JPEG with lower quality
        img_byte_arr = io.BytesIO()
        img.convert('RGB').save(img_byte_arr, format='JPEG', quality=60)  # Lower quality JPEG
        img_byte_arr.seek(0)

        # Use reportlab to create a PDF from the image
        packet = io.BytesIO()
        c = canvas.Canvas(packet, pagesize=(w, h))  # No scaling here since we're using lower DPI

        # Draw the image
        c.drawImage(ImageReader(img_byte_arr), 0, 0, width=w, height=h)
        
        # Set up text overlay
        c.saveState()
        c.setFillColorRGB(0, 0, 0, alpha=0.01)  # Almost transparent black for overlay
        c.setFont("DejaVuSans", 10)  # Adjust font size as per need

        # Draw text based on bounding boxes
        for i, line in enumerate(data['text']):
            if line.strip():
                left = data['left'][i]
                top = data['top'][i]
                width = data['width'][i]
                height = data['height'][i]
                
                # Position text
                c.drawString(left, h - top - height, line)

        c.restoreState()
        c.save()

        packet.seek(0)
        
        # Create a new PDF with reportlab's output
        new_pdf = PdfReader(packet)
        page = new_pdf.pages[0]
    
        merger.add_page(page)

    with open(output_path, "wb") as out:
        merger.write(out)

    print(f"Searchable and selectable PDF saved to {output_path}")


def process_directory(directory):
    non_searchable_dir = os.path.join(directory, "non-searchable")
    if not os.path.exists(non_searchable_dir):
        os.makedirs(non_searchable_dir)

    pdf_files = [f for f in os.listdir(directory) if f.lower().endswith('.pdf')]
    total_files = len(pdf_files)
    print(f"Found {total_files} PDF files in {directory} to process.")
    item_number = 0
    
    for index, pdf_file in enumerate(pdf_files, start=1):
        item_number = item_number + 1
        # print(f"---- ({item_number}/{total_files}) ----")
        pdf_path = os.path.join(directory, pdf_file)
        if not is_pdf_searchable(pdf_path):
            # First, copy the original PDF to the non-searchable directory
            new_name = os.path.splitext(pdf_file)[0] + "_ns.pdf"
            new_path = os.path.join(non_searchable_dir, new_name)
            shutil.copy(pdf_path, new_path)
            print(f"({item_number}/{total_files}) Copied {pdf_file} to {new_path}")
            
            # Then, process the PDF to make it searchable, overwriting the original
            process_single_pdf(pdf_path, directory)
        else:
            print(f"({item_number}/{total_files}) File {pdf_file} is already searchable. Skipping.")

def main():
    if len(sys.argv) != 2:
        print("Usage: python woo-ocrpdf.py <pdf_file_or_directory>")
        sys.exit(1)

    path = sys.argv[1]
    if os.path.isfile(path):
        if path.lower().endswith('.pdf'):
            directory = os.path.dirname(path)
            non_searchable_dir = os.path.join(directory, "non-searchable")
            if not os.path.exists(non_searchable_dir):
                os.makedirs(non_searchable_dir)
            
            if not is_pdf_searchable(path):
                # First, copy the original PDF to the non-searchable directory
                new_name = os.path.splitext(path)[0] + "_ns.pdf"
                new_path = os.path.join(non_searchable_dir, new_name)
                shutil.copy(path, new_path)
                print(f"Copied {path} to {new_path}")
                
                # Then, process the PDF to make it searchable, overwriting the original
                process_single_pdf(path, directory)
            else:
                print(f"File {os.path.basename(path)} is already searchable. Skipping.")
        else:
            print("The file provided is not a PDF.")
            
    elif os.path.isdir(path):
        process_directory(path)
        
    else:
        print("Invalid path provided. It must be either a PDF file or a directory.")

if __name__ == "__main__":
    main()
