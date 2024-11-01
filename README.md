Follow me on X: https://x.com/leon1969

These scripts are used in the following order:
1. woo-extract-docrn.py: analyse large PDF files with multiple embedded documents identifiable through a documentnumber in the top-right, bottom-right, top-left and bottom-left corner. The result of this script is a file with per line <file name> <doc nr> <page-range>. The file is equal to the name of the analysed PDF file, with added "*_document_numbers.txt".
2. woo-extract.py: uses the file created in step 1 to create separate PDF files from the different embedded documents in the PDF analysed in step 1.
3. woo-ocrpdf.py: OCR's a non searchable PDF. Takes as input parameter a PDF file or a folder containing PDF's. It copies non-searchable PDF's to a subfolder called "non-searchable" and saves the created searchable PDF at the oroginal file location.

