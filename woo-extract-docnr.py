import sys
import fitz  # PyMuPDF
from PIL import Image
import pytesseract
import re
import itertools
from operator import itemgetter
import time

def extract_document_number(page, page_num, total_pages, start_time):
    """
    Extracts the document number from the upper right corner of a PDF page.

    :param page: A fitz.Page object from PyMuPDF
    :param page_num: The number of the current page
    :param total_pages: Total number of pages in the document
    :param start_time: Start time for time estimation
    :return: Extracted document number or None if not found
    """
    # Render page at a reasonable resolution
    pix = page.get_pixmap(matrix=fitz.Matrix(300/72, 300/72))
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

    # Crop the upper right corner where the document number is expected to be
    width, height = img.size
    cropped = img.crop((width - 500, 0, width, 400))  # Adjust these values based on your PDF layout

    # Perform OCR on the cropped image
    text = pytesseract.image_to_string(cropped)
    # Use regex to find a number or document number pattern
    match = re.search(r'\b\d{1,}\b', text)  # Adjust regex as needed for your document number format
    doc_number = match.group() if match else None
    
    # Time estimation
    elapsed_time = time.time() - start_time
    avg_time_per_page = elapsed_time / page_num
    remaining_time = avg_time_per_page * (total_pages - page_num)
    
    print(f"Processing page {page_num}/{total_pages}: OCR'd document number is {doc_number if doc_number else 'not found'} "
          f"| Estimated remaining time: {remaining_time:.0f} seconds")
    return doc_number

def pages_to_ranges(pages):
    """
    Convert a list of page numbers into a string where consecutive numbers are replaced by a range.

    :param pages: List of page numbers
    :return: String of page ranges
    """
    if not pages: return ""
    ranges = []
    for k, g in itertools.groupby(enumerate(pages), lambda x: x[0] - x[1]):
        group = list(map(itemgetter(1), g))
        if len(group) == 1:
            ranges.append(str(group[0]))
        else:
            ranges.append(f"{group[0]}-{group[-1]}")
    return ", ".join(ranges)

def process_pdf(input_pdf):
    """
    Processes a PDF file, OCRs every page, and records document numbers with their page numbers.

    :param input_pdf: Path to the PDF file
    :return: None, writes information to a text file
    """
    doc = fitz.open(input_pdf)
    total_pages = len(doc)
    start_time = time.time()
    
    print(f"Starting OCR on '{input_pdf}', total pages: {total_pages}")
    
    output_file = f"{input_pdf.split('.')[0]}_document_numbers.txt"
    
    with open(output_file, 'w', encoding='utf-8') as file:
        doc_numbers = {}
        for page_num, page in enumerate(doc, start=1):
            doc_number = extract_document_number(page, page_num, total_pages, start_time)
            if doc_number:
                if doc_number in doc_numbers:
                    doc_numbers[doc_number].append(page_num)
                else:
                    doc_numbers[doc_number] = [page_num]
        
        for doc_number, pages in doc_numbers.items():
            file.write(f"{input_pdf} {doc_number} {pages_to_ranges(pages)}\n")
    
    total_time = time.time() - start_time
    print(f"Finished processing. Time taken: {total_time:.2f} seconds.")
    print(f"Results saved to {output_file}")

# Main execution
if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python woo-extract-docnr.py <input_pdf>")
        print("Creates <input_pdf>_document_numbers.txt which has OCR'd enclosed document numbers with their page-range.")
        sys.exit(1)

    input_pdf = sys.argv[1]
    process_pdf(input_pdf)
