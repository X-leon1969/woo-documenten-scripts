import sys
from PyPDF2 import PdfReader, PdfWriter

def split_pdf(input_pdf, document_number, page_range):
    """
    Extracts specific pages from a PDF and saves them as a new PDF.

    :param input_pdf: Path to the input PDF
    :param document_number: The number to use in the output filename
    :param page_range: String describing the pages to extract
    :return: None, saves a new PDF file
    """
    reader = PdfReader(input_pdf)
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

    for page in pages:
        if page <= len(reader.pages):
            writer.add_page(reader.pages[page - 1])  # PyPDF2 uses 0-indexing

    # Create output filename
    base_name = input_pdf.split('.')[0]
    output_file = f"{document_number} {page_range.replace(',', '_')}__ {base_name}.pdf"
    
    with open(output_file, "wb") as output_stream:
        writer.write(output_stream)
    print(f"Created new PDF: {output_file}")

# Main execution
if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("Usage: python woo-extract.py <input_pdf> <document_number> <page_ranges>")
        sys.exit(1)

    input_pdf, document_number, page_range = sys.argv[1], sys.argv[2], sys.argv[3]
    split_pdf(input_pdf, document_number, page_range)
