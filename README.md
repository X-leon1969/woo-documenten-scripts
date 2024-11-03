Follow me on X: https://x.com/leon1969

When reading and evaluating the FOIA'd documents from the Dutch Ministry of Health some FOIA'd PDF's contained a lot of embedded documents with some of them not searchable. In order to be able to correctly analyse these documents, these scripts are used to savethese different documents as seperate PDF files and make sure they are text searchable.

These scripts are used in the following order:
1. woo-extract-docrn.py: analyse large PDF files with multiple embedded documents identifiable through a documentnumber in the top-right, bottom-right, top-left or bottom-left corner. The result of this script is a file with per line "filename docnr page-range". The file is equal to the name of the analysed PDF file, with added "*_document_numbers.txt".
2. woo-extract.py: uses the file created in step 1 to create separate PDF files from the different embedded documents in the PDF analysed in step 1.
3. woo-ocrpdf.py: OCR's a non searchable PDF. Takes as input parameter a PDF file or a folder containing PDF's. It copies non-searchable PDF's to an underlying subfolder called "non-searchable" and saves the created searchable PDF at the original file location.
4. woo-datespec.config: config for retrieving the date of a document
5. woo-datespec.py: retrieves the document date from its first page
6. woo-getupdates.py: spider open.minvws.nl "besluiten" search, and download all "besluiten". Save meta data into excel file. Download all "inventaris" files.
