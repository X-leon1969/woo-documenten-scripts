import sys
from PyPDF2 import PdfReader, PdfWriter
import os
import time
import glob

def parse_line(line):
    """
    Parse each line from the input file into components needed for splitting.
    
    :param line: A string containing 'pdf_path document_number page_range'
    :return: Tuple of (pdf_path, document_number, page_range)
    """
    parts = line.strip().split(' ')
    if len(parts) != 3:
        raise ValueError(f"Incorrect format in line: {line}")
    return parts[0], parts[1], parts[2]

def process_pdf(pdf_path, document_number, page_range, reader):
    """
    Extracts specified pages from a PDF and saves them as a new PDF.

    :param pdf_path: Path to the input PDF
    :param document_number: The number to use in the output filename
    :param page_range: String describing the pages to extract
    :param reader: The PdfReader object for the PDF
    :return: None, saves a new PDF file
    """
    total_pages = len(reader.pages)
    writer = PdfWriter()

    # Parse page range
    pages = []
    for part in page_range.split(','):
        if '-' in part:
            a, b = map(int, part.split('-'))
            pages.extend(range(a, b + 1))
        else:
            pages.append(int(part))
    pages = sorted(list(set(pages)))  # Ensure no duplicates and in order

    print(f"Extracting pages {page_range} from {os.path.basename(pdf_path)}")

    for i, page in enumerate(pages, 1):
        if page <= total_pages:
            writer.add_page(reader.pages[page - 1])  # PyPDF2 uses 0-indexing
            # print(f"Extracted page {page}")

    # Create output filename
    base_name = os.path.basename(pdf_path).split('.')[0]  # Strip path information
    output_file = f"{document_number} p{page_range.replace(',', '_')} __ {base_name}.pdf"
    
    with open(output_file, "wb") as output_stream:
        writer.write(output_stream)
    
    print(f"New PDF saved as: {output_file}")

def split_pdfs_from_file(instructions_file_pattern):
    """
    Process files containing lines of PDF split instructions, allowing wildcard patterns.
    
    :param instructions_file_pattern: Path pattern to the file(s) containing split instructions
    """
    for instructions_file in glob.glob(instructions_file_pattern):
        current_pdf = None
        reader = None

        with open(instructions_file, 'r') as file:
            for line in file:
                try:
                    pdf_path, document_number, page_range = parse_line(line)
                    
                    if current_pdf != pdf_path:
                        if reader:
                            reader.stream.close()  # Close the previous PDF's stream
                        reader = PdfReader(pdf_path)
                        current_pdf = pdf_path
                    
                    start_time = time.time()
                    process_pdf(pdf_path, document_number, page_range, reader)
                    end_time = time.time()
                    print(f"Time taken for this operation: {end_time - start_time:.2f} seconds.")
                    print("-" * 50)  # Visual separator for each operation

                except Exception as e:
                    print(f"Error processing line '{line.strip()}' from {instructions_file}: {str(e)}")
                    continue

        if reader:
            reader.stream.close()  # Ensure the last opened PDF is closed

# Main execution
if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python woo-extract.py <instructions_file_pattern>")
        print("Input file has per line <pdf_name> <document_id> <page_range>.")
        print("Output will be extracted pdf's per <document_id>.")
        sys.exit(1)

    instructions_file_pattern = sys.argv[1]
    split_pdfs_from_file(instructions_file_pattern)
