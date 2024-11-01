import sys
import fitz  # PyMuPDF
from PIL import Image
import pytesseract
import re
import itertools
from operator import itemgetter
import time
import os
import re
import numpy as np
import cv2
from threading import Timer

# Define timeout handler
def timeout_handler():
    print("Script execution has exceeded the time limit. Aborting.")
    sys.exit(1)

# Set up timeout
class Timeout:
    def __init__(self, seconds=1, error_message='Timeout'):
        self.seconds = seconds
        self.error_message = error_message
    def handle_timeout(self, signum, frame):
        raise TimeoutError(self.error_message)
    def __enter__(self):
        self.timer = Timer(self.seconds, timeout_handler)
        self.timer.start()
    def __exit__(self, type, value, traceback):
        self.timer.cancel()

def extract_document_number(page, page_num, total_pages, corner):
    """
    Extracts the document number from the specified corner of a PDF page.

    :param page: A fitz.Page object from PyMuPDF
    :param page_num: The number of the current page
    :param total_pages: Total number of pages in the document
    :param corner: The corner where the document number is located
    :return: Extracted document number or None if not found
    """
    # Render page at a reasonable resolution
    pix = page.get_pixmap(matrix=fitz.Matrix(300/72, 300/72))
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

    # Define cropping area based on corner
    width, height = img.size
    boxsize2 = 500
    boxsize1 = 500
    x1, y1, x2, y2 = {
        "top-left": (0, 0, boxsize1, boxsize2),
        "top-right": (width - boxsize1, 0, width, boxsize2),
        "bottom-left": (0, height - boxsize2, boxsize1, height),
        "bottom-right": (width - boxsize1, height - boxsize2, width, height)
    }.get(corner, (width - boxsize1, 0, width, boxsize2))  # Default to top-right if invalid corner

    cropped = img.crop((x1, y1, x2, y2))

    # Perform OCR on the cropped image
    text = pytesseract.image_to_string(cropped)
    # Use regex to find a number or document number pattern
    match = re.search(r'\b\d{1,}\b', text)  # Adjust regex as needed
    doc_number = match.group() if match else None
    
    if not doc_number:
        # def extract_document_number_from_red_box(input_pdf, page_number=0):
        """
        Extracts document number from a red box in the top-right corner of a PDF page.

        :param input_pdf: Path to the PDF file
        :param page_number: Page number to extract from (default is 0 for the first page)
        :return: Extracted document number or None if not found
        """
        # Open the PDF file
        # document = fitz.open(input_pdf)
        
        # if page_number >= len(document):
            # print(f"Error: Page {page_number} not found in {input_pdf}")
            # return None

        # Get the page
        # page = document[page_number]
        # pix = page.get_pixmap(matrix=fitz.Matrix(300/72, 300/72))  # Render at 300 DPI for better OCR
        # img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        
        # Convert PIL Image to numpy array for OpenCV
        np_img = np.array(img)

        # Define the region of interest (ROI) - top-right corner
        height, width = np_img.shape[:2]
        roi = np_img[0:int(height/4), int(width*3/4):width]  # You might need to adjust these based on your PDF's layout

        # Convert ROI to HSV color space for easier color detection
        hsv_roi = cv2.cvtColor(roi, cv2.COLOR_RGB2HSV)

        # Define range for red color in HSV
        lower_red = np.array([0, 50, 50])
        upper_red = np.array([10, 255, 255])

        # Create mask for red color
        mask = cv2.inRange(hsv_roi, lower_red, upper_red)

        # Find contours in the mask
        contours, _ = cv2.findContours(mask, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)

        # Look for the largest contour which might be the red box
        if contours:
            largest_contour = max(contours, key=cv2.contourArea)
            
            # Get the bounding rectangle
            x, y, w, h = cv2.boundingRect(largest_contour)
            
            # Crop the region inside the red box
            red_box = roi[y:y+h, x:x+w]

            # Convert the cropped area back to PIL Image for OCR
            red_box_pil = Image.fromarray(red_box)
            
            # Use Tesseract to recognize text
            text = pytesseract.image_to_string(red_box_pil, config='--psm 6')
            
            # Use regex to extract what looks like a document number
            # import re
            match = re.search(r'\b\d+\b', text)  # Assuming the document number is a sequence of digits
            if match:
                doc_number = match
            # else:
                # print("No document number found in the red box.")
        else:
            print("No red box detected on the page.")
            
    
    # Time estimation
    elapsed_time = time.time() - start_time
    avg_time_per_page = elapsed_time / page_num if page_num > 0 else 0
    remaining_time = avg_time_per_page * (total_pages - page_num + 1)
    
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

def process_pdf(input_pdf, corner):
    """
    Processes a PDF file, OCRs every page, and records document numbers with their page numbers.

    :param input_pdf: Path to the PDF file
    :param corner: The corner where document numbers are expected to be found
    :return: None, writes information to a text file
    """
    global start_time
    start_time = time.time()
    
    doc = fitz.open(input_pdf)
    total_pages = len(doc)
    
    print(f"Starting OCR on '{input_pdf}', total pages: {total_pages}")
    
    output_file = f"{input_pdf.split('.')[0]}_document_numbers.txt"
    
    with open(output_file, 'w', encoding='utf-8') as file:
        doc_numbers = {}
        with Timeout(3600):  # 3600 seconds timeout
            for page_num, page in enumerate(doc, start=1):
                doc_number = extract_document_number(page, page_num, total_pages, corner)
                if doc_number:
                    if doc_number in doc_numbers:
                        doc_numbers[doc_number].append(page_num)
                    else:
                        doc_numbers[doc_number] = [page_num]
        
        for doc_number, pages in doc_numbers.items():
            file.write(f"{os.path.basename(input_pdf)} {doc_number} {pages_to_ranges(pages)}\n")
    
    total_time = time.time() - start_time
    print(f"Finished processing. Time taken: {total_time:.2f} seconds.")
    print(f"Results saved to {output_file}")

# Main execution
if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python script.py <input_pdf> <corner>")
        print("Example corners: 'top-left', 'top-right', 'bottom-left', 'bottom-right'")
        sys.exit(1)

    input_pdf, corner = sys.argv[1], sys.argv[2]
    if corner not in ["top-left", "top-right", "bottom-left", "bottom-right"]:
        print("Invalid corner specified.")
        sys.exit(1)

    try:
        process_pdf(input_pdf, corner)
    except TimeoutError:
        print("Script execution has exceeded the time limit. Aborting.")
