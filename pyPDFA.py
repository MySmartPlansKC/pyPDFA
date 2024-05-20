import pdfrw
import logging
import os
import shutil
import subprocess
import sys
import time

# Versioning
__version__ = "1.3.0"
# pyinstaller --onefile --name pyPDFA-V1.3.0 pyPDFA.py

# Configure logging
logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s',
                    datefmt='%Y-%m-%d %H:%M:%S',
                    handlers=[
                        logging.FileHandler("pdf_conversion.log"),
                        logging.StreamHandler()
                    ])


def safe_remove(path):
    try:
        os.unlink(path)
    except PermissionError:
        logging.error(f"Could not delete {path} - it may be in use.")


def safe_rmtree(directory):
    for root, dirs, files in os.walk(directory, topdown=False):
        for name in files:
            file_path = os.path.join(root, name)
            safe_remove(file_path)
        for name in dirs:
            dir_path = os.path.join(root, name)
            try:
                os.rmdir(dir_path)
            except OSError as e:
                logging.error(f"Could not delete directory {dir_path}: {str(e)}")


def get_base_path():
    if getattr(sys, 'frozen', False):
        # If the application is frozen using PyInstaller
        return os.path.dirname(sys.executable)
    else:
        # Normal execution (e.g., script or interactive)
        return os.path.dirname(os.path.abspath(__file__))


def get_pdf_page_count(pdf_path):
    try:
        reader = pdfrw.PdfReader(pdf_path)
        page_count = len(reader.pages)
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
            "-dQUIET",
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

        result = subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        if result.returncode == 0:
            logging.info(f"Successfully converted {source_path} to PDF/A.")
            return True, page_count
        else:
            logging.error(f"Ghostscript error: {result.stderr.decode()}")
            return False, page_count
    except subprocess.CalledProcessError as e:
        logging.error(f"Failed to convert {source_path}: {str(e)}")
        return False, 0


def check_and_clear_directory(directory):
    if os.path.exists(directory):
        if os.listdir(directory):  # Check if the directory is not empty
            response = input(
                f"Directory {directory} is not empty. Delete all contents? (y/n): "
            )
            if response.lower() == 'y':
                safe_rmtree(directory)  # Remove the directory and its contents, safely
                os.makedirs(directory, exist_ok=True)  # Recreate the empty directory
                logging.info(f"All contents of {directory} have been deleted.")
                time.sleep(1)  # Wait for a moment after deleting the contents
            else:
                logging.error("Operation aborted by the user.")
                return False
    else:
        os.makedirs(directory, exist_ok=True)
    return True


def remove_empty_directories(path, root_dir):
    try:
        if os.path.isdir(path) and not os.listdir(path):
            os.rmdir(path)
            logging.info(f"Removed empty directory: {path}")
            parent_dir = os.path.dirname(path)
            if parent_dir != root_dir:  # Ensure the root directory is not removed
                remove_empty_directories(parent_dir, root_dir)
    except OSError as e:
        logging.error(f"Error removing directory {path}: {str(e)}")


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
    for root, dirs, files in os.walk(input_dir):
        for file in files:
            if file.endswith('.pdf'):
                source_file = os.path.join(root, file)
                relative_path = os.path.relpath(root, input_dir)
                output_file_dir = os.path.join(output_dir, relative_path)
                output_file = os.path.join(output_file_dir, file)
                error_file_dir = os.path.join(error_dir, relative_path)

                os.makedirs(output_file_dir, exist_ok=True)
                os.makedirs(error_file_dir, exist_ok=True)

                success, _ = convert_to_pdfa(source_file, output_file)
                if success:
                    os.remove(source_file)
                    remove_empty_directories(os.path.dirname(source_file), input_dir)
                else:
                    shutil.move(source_file, os.path.join(error_file_dir, file))
                    logging.warning(f"File moved to error directory: {file}")
                    has_errors = True

    if has_errors:
        logging.info("There has been at least one error, please check the PDF_Not_Converted folder.")


if __name__ == '__main__':
    base_path = get_base_path()

    # Testing paths
    # input_directory = os.path.join(base_path, "..", "xPDFTestFiles", "PDFA_IN")
    # output_directory = os.path.join(base_path, "..", "xPDFTestFiles", "PDFA_OUT")
    # error_directory = os.path.join(base_path, "..", "xPDFTestFiles", "PDF_Not_Converted")

    # Uncomment below for production paths
    input_directory = os.path.join(base_path, "PDFA_IN")
    output_directory = os.path.join(base_path, "PDFA_OUT")
    error_directory = os.path.join(base_path, "PDF_Not_Converted")

    logging.info(f"Starting PDFA Conversion v{__version__}")
    time.sleep(1)

    os.makedirs(output_directory, exist_ok=True)
    os.makedirs(error_directory, exist_ok=True)

    batch_convert(input_directory, output_directory, error_directory)
    logging.info("Batch conversion process completed.")