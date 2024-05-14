# import fitz
import logging
import os
import shutil
import subprocess
import sys
import time

# Versioning
__version__ = "1.1.1"
# pyinstaller --onefile --name pyPDFA-V1.1.1 pyPDFA.py

# Configure logging
logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s',
                    datefmt='%Y-%m-%d %H:%M:%S',
                    handlers=[
                        logging.FileHandler("pdf_conversion.log"),
                        logging.StreamHandler()
                    ])


def get_base_path():
    if getattr(sys, 'frozen', False):
        # If the application is frozen using PyInstaller
        return os.path.dirname(sys.executable)
    else:
        # Normal execution (e.g., script or interactive)
        return os.path.dirname(os.path.abspath(__file__))


def get_pdf_page_count(pdf_path):
    try:
        doc = fitz.open(pdf_path)
        page_count = len(doc)
        doc.close()
        return page_count
    except Exception as e:
        logging.error(f"Error getting page count for {pdf_path}: {e}")
        return 0


def convert_to_pdfa(source_path, output_path):
    try:
        gs_executable = r'C:\Program Files\gs\gs10.03.0\bin\gswin64c.exe'
        page_count = get_pdf_page_count(source_path)
        cmd = [
            gs_executable,
            "-dQUIET"
            "-dPDFA",
            "-dBATCH",
            "-dNOPAUSE",
            "-dNOOUTERSAVE",
            "-sDEVICE=pdfwrite",
            "-sProcessColorModel=DeviceRGB",
            "-sPDFACompatibilityPolicy=1",
            f"-sOutputFile={output_path}",
            source_path
        ]
        for page_num in range(1, page_count + 1):
            logging.info(f"Processing page {page_num} of {page_count}")

        subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        logging.info(f"Successfully converted {source_path} to PDF/A.")
        return True, page_count
    except subprocess.CalledProcessError as e:
        logging.error(f"Failed to convert {source_path}: {e}")
        return False


def check_and_clear_directory(directory):
    if os.path.exists(directory):
        if os.listdir(directory):  # Check if the directory is not empty
            response = input(
                f"Directory {directory} is not empty. Delete all contents? (y/n): "
            )
            if response.lower() == 'y':
                shutil.rmtree(directory)  # Remove the directory and its contents
                os.makedirs(directory)  # Recreate the empty directory
                logging.info(f"All contents of {directory} have been deleted.")
                time.sleep(1)  # Wait for a moment after deleting the contents
            else:
                logging.error("Operation aborted by the user.")
                return False
    else:
        os.makedirs(directory, exist_ok=True)
    return True


def batch_convert(input_dir, output_dir, error_dir):

    # Check and possibly clear the output directory
    if not check_and_clear_directory(output_directory):
        return
    time.sleep(1)

    # Check and possibly clear the error directory
    if not check_and_clear_directory(error_directory):
        return
    time.sleep(1)

    if not os.path.exists(input_dir) or not os.listdir(input_dir):
        logging.error(f"No files to process in {input_dir}.")
        return

    has_errors = False
    for filename in os.listdir(input_dir):
        if filename.endswith('.pdf'):
            source_file = os.path.join(input_dir, filename)
            output_file = os.path.join(output_dir, filename)
            if convert_to_pdfa(source_file, output_file):
                # Comment out the line below for testing
                os.remove(source_file)
                # logging.info(f"Conversion successful, skipping deletion for testing: {source_file}")
            else:
                error_destination = os.path.join(error_dir, filename)
                # Comment out the lines below for testing
                shutil.move(source_file, error_destination)
                logging.warning(f"File moved to error directory: {filename}")
                # logging.warning(f"Conversion failed, skipping move to error directory for testing: {filename}")
                has_errors = True

    if has_errors:
        logging.info("There has been at least one error, please check the PDF_Not_Converted folder.")


if __name__ == '__main__':
    base_path = get_base_path()

    # input_directory = r"..\xPDFTestFiles"
    input_directory = os.path.join(base_path, 'PDFA_IN')
    output_directory = os.path.join(base_path, 'PDFA_OUT')
    error_directory = os.path.join(base_path, 'PDF_Not_Converted')

    logging.info(f"Starting PDFA Conversion v{__version__}")
    time.sleep(1)

    os.makedirs(output_directory, exist_ok=True)
    os.makedirs(error_directory, exist_ok=True)

    batch_convert(input_directory, output_directory, error_directory)
    logging.info("Batch conversion process completed.")
