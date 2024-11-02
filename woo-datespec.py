import os
import re
import sys
import glob
import logging
from PyPDF2 import PdfReader
from datetime import datetime
import dateutil.parser
from configparser import ConfigParser, NoSectionError, NoOptionError
import pytz

def get_script_dir():
    return os.path.dirname(os.path.abspath(__file__))

# Setup logging
script_dir = get_script_dir()
logging_file = os.path.join(script_dir, os.path.basename(sys.argv[0]).replace('.py', '.log'))
logging.basicConfig(filename=logging_file, level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

def read_config():
    script_dir = get_script_dir()
    config_file = os.path.join(script_dir, os.path.basename(sys.argv[0]).replace('.py', '.config'))
    
    config = ConfigParser()
    try:
        if not config.read(config_file):
            raise FileNotFoundError(f"Configuration file not found at {config_file}. Please ensure the config file exists in the same directory as the script.")
        
        date_formats = config.get('DateFormats', 'DATE_FORMATS').split(',')
        date_identifiers = config.get('DateIdentifiers', 'DATE_IDENTIFIERS').split(',')
        languages = config.get('Languages', 'LANGUAGES').split(',')
        search_on_next_line_after = config.get('DateSearchRules', 'SEARCH_ON_NEXT_LINE_AFTER', fallback='Datum').lower()
        search_subfolders = config.getboolean('ProcessingRules', 'SEARCH_SUBFOLDERS', fallback=True)
        redo = config.getboolean('ProcessingRules', 'REDO', fallback=False)
        allowed_years = set(int(year) for year in config.get('DateValidation', 'ALLOWED_YEARS', fallback='').split(','))
        
        return date_formats, date_identifiers, languages, search_on_next_line_after, search_subfolders, allowed_years, redo

    except FileNotFoundError as e:
        print(f"Error: {e}")
        print("Help: The configuration file must be in the same directory as the script with the same name but with a '.config' extension.")
        sys.exit(1)

    except NoSectionError as e:
        print(f"Error: {e}")
        print("Help: Ensure all required sections are present in the configuration file.")
        print("The file should have sections: DateFormats, DateIdentifiers, DateSearchRules, ProcessingRules, and DateValidation.")
        sys.exit(1)

    except NoOptionError as e:
        print(f"Error: {e}")
        print("Help: Make sure you have all required options under each section in the config file.")
        # Here you might want to list the expected options for each section, but for brevity:
        print("Check that all necessary options are correctly spelled and present.")
        sys.exit(1)

    except ValueError as e:
        print(f"Error: {e}")
        print("Help: The 'ALLOWED_YEARS' in the 'DateValidation' section should contain only numbers separated by commas.")
        sys.exit(1)

    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        print("Please check the configuration file for any typos or formatting issues.")
        sys.exit(1)

def is_valid_date(date_string):
    try:
        datetime.strptime(date_string, "%Y%m%d")
        return True
    except ValueError:
        return False

def export_text(pdf_path, text):
    txt_path = pdf_path.rsplit('.', 1)[0] + '.txt'
    with open(txt_path, 'w', encoding='utf-8') as txt_file:
        txt_file.write(text)
    logging.info(f"Exported text to {txt_path}")

def extract_date_from_text(text, date_formats, date_identifiers, search_on_next_line_after, allowed_years):
    tzinfos = {"CEST": 3600, "JEN": 3600, "IEE": 3600 }  # CEST is +1 hour from UTC, hence 3600 seconds

    lines = text.split('\n')
    found_dates = []

    for i, line in enumerate(lines):
        for identifier in date_identifiers:
            if identifier in line:
                logging.info(f"Identifier '{identifier}' found on line: {line}")
                if i + 2 < len(lines):  # Check if we can look at two more lines
                    potential_day = lines[i + 1].strip()
                    potential_date = lines[i + 2].strip()
                    combined_date_str = f"{potential_day} {potential_date}"
                    try:
                        parsed_date = dateutil.parser.parse(combined_date_str, fuzzy=True, tzinfos=tzinfos)
                        parsed_date = parsed_date.replace(tzinfo=None) if parsed_date.tzinfo else parsed_date  # Make naive if it's timezone-aware
                        if parsed_date.year in allowed_years:
                            found_dates.append(parsed_date)
                        else:
                            logging.info(f"Date {parsed_date.strftime('%Y%m%d')} discarded due to invalid year.")
                    except ValueError:
                        pass  # No logging for unsuccessful parsing

                date_str = line.split(identifier)[-1].strip()
                if identifier.lower() == search_on_next_line_after:
                    date_str = lines[i + 1].strip() if i + 1 < len(lines) else ''
                try:
                    parsed_date = dateutil.parser.parse(date_str, fuzzy=True, tzinfos=tzinfos)
                    parsed_date = parsed_date.replace(tzinfo=None) if parsed_date.tzinfo else parsed_date  # Make naive if it's timezone-aware
                    if parsed_date.year in allowed_years:
                        found_dates.append(parsed_date)
                    else:
                        logging.info(f"Date {parsed_date.strftime('%Y%m%d')} discarded due to invalid year.")
                except ValueError:
                    pass  # No logging for unsuccessful parsing

    # If multiple dates found, return the first one
    if found_dates:
        return min(found_dates, key=lambda x: x).strftime('%Y%m%d')
    logging.info("No valid date found.")
    return None
    
def rename_pdf(pdf_path, new_prefix, old_filename=None):
    directory, filename = os.path.split(pdf_path)
    if old_filename is None:
        old_filename = filename
    new_filename = f"{new_prefix} {old_filename}"
    new_pdf_path = os.path.join(directory, new_filename)
    os.rename(pdf_path, new_pdf_path)
    logging.info(f"Renamed: {filename} -> {new_filename}")
    print(f"Renamed: {filename}\n         {new_filename}")

def process_pdf(pdf_path, date_formats, date_identifiers, languages, search_on_next_line_after, allowed_years, redo, index, total):
    filename = os.path.basename(pdf_path)
    logging.info(f"Processing PDF {index} of {total}: {filename}")
    print(f"Processing PDF {index} of {total}: {filename}")
    date_prefix = filename[:8]
    
    if is_valid_date(date_prefix) and not redo:  # already prefixed
        return

    reader = PdfReader(pdf_path)
    if len(reader.pages) > 0:
        first_page = reader.pages[0]
        text = first_page.extract_text()
        export_text(pdf_path, text)
        
        date_found = extract_date_from_text(text, date_formats, date_identifiers, search_on_next_line_after, allowed_years)
        if date_found:
            if redo:
                original_filename = re.sub(r'^\d{8} ', '', filename)
                rename_pdf(pdf_path, date_found, original_filename)
            else:
                rename_pdf(pdf_path, date_found)
        else:
            if not filename.startswith("UNKNOWN_"):
                rename_pdf(pdf_path, "UNKNOWN_", filename)   
                
def main(target):
    date_formats, date_identifiers, languages, search_on_next_line_after, search_subfolders, allowed_years, redo = read_config()
    
    pdf_files = []
    if '*' in target or '?' in target:
        pdf_files = glob.glob(target)
        if not pdf_files:
            print(f"No PDF files matched the wildcard '{target}'.")
            logging.error(f"No PDF files matched the wildcard '{target}'.")
            sys.exit(1)
    elif os.path.isdir(target):
        if search_subfolders:
            pdf_files = [os.path.join(root, file) for root, _, files in os.walk(target) for file in files if file.lower().endswith('.pdf')]
        else:
            pdf_files = [os.path.join(target, file) for file in os.listdir(target) if file.lower().endswith('.pdf')]
    elif os.path.isfile(target) and target.lower().endswith('.pdf'):
        pdf_files = [target]
    else:
        print("Error: The provided path does not match any PDF file or directory.")
        logging.error(f"Error: The provided path '{target}' does not match any PDF file or directory.")
        sys.exit(1)

    total_files = len(pdf_files)
    print(f"Found {total_files} PDF files to process.")
    for index, pdf_path in enumerate(pdf_files, 1):
        process_pdf(pdf_path, date_formats, date_identifiers, languages, search_on_next_line_after, allowed_years, redo, index, total_files)

def show_help():
    print("Help for PDF Date Extraction and Renaming Script:")
    print("\nUsage:")
    print("  python woo-datespec.py <pdf_file_or_directory_or_wildcard>")
    print("\nOptions:")
    print("  --help, -h   Show this help message and exit.")
    print("\nDescription:")
    print("  This script searches for dates within PDF files based on specified identifiers in a configuration file.")
    print("  It renames the PDF files by adding a date at the beginning of the filename if found. If no date is found,")
    print("  'UNKNOWN_' is prefixed to the filename. The script can handle PDF files directly, directories, or wildcard patterns.")
    
    print("\nKey Features:")
    print("  - Date Extraction: Looks for dates after identifiers or on the next line after specific identifiers.")
    print("  - Date Validation: Ensures dates are within allowed years specified in the configuration.")
    print("  - Redo Option: When enabled, forces re-evaluation of all files, potentially renaming them even if previously named with a date.")
    print("  - Subfolder Search: Optionally searches for PDFs in subdirectories.")
    print("  - Text Extraction: Extracts and saves the first page's text of each PDF to a .txt file.")
    print("  - Logging: Logs operations and errors to a file named after the script with a .log extension.")
    
    print("\nConfiguration:")
    print("  - The script uses a configuration file named 'woo-datespec.config' in the same directory.")
    print("  - Contains sections for DateFormats, DateIdentifiers, DateSearchRules, ProcessingRules, and DateValidation.")
    print("  - Customize the config to adjust date formats, identifiers, timezone handling (e.g., CEST, JEN, IEE) and processing rules.")
    
    print("\nTroubleshooting:")
    print("  - If the script can't find your configuration file, ensure it's named correctly and in the script's directory.")
    print("  - For timezone issues, review the 'tzinfos' dictionary in the script for correct offsets.")
    print("  - If dates are not recognized, check if the date format in the PDF matches those listed in the config.")
    
    print("\nLogging:")
    print(f"  - Errors and information are logged to '{os.path.basename(sys.argv[0]).replace('.py', '.log')}' in the script's directory.")
    
    print("\nExit with Ctrl+C to interrupt the script.")
    
if __name__ == "__main__":
    if len(sys.argv) < 2 or sys.argv[1] in ['--help', '-h']:
        show_help()
        sys.exit(0)
    
    if len(sys.argv) != 2:
        show_help()
        logging.error("Script usage error: Missing argument for PDF file, directory, or wildcard path.")
        sys.exit(1)
    
    target = sys.argv[1]
    main(target)
