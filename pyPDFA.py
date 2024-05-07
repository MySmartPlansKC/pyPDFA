import sys
import os
import shutil
import subprocess
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s',
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


def convert_to_pdfa(source_path, output_path):
    try:
        gs_executable = r'C:\Program Files\gs\gs10.03.0\bin\gswin64c.exe'
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
        subprocess.run(cmd, check=True)
        logging.info(f"Successfully converted {source_path} to PDF/A.")
        return True
    except subprocess.CalledProcessError as e:
        logging.error(f"Failed to convert {source_path}: {e}")
        return False


def batch_convert(input_dir, output_dir, error_dir):
    if not os.path.exists(input_dir) or not os.listdir(input_dir):
        logging.error(f"No files to process in {input_dir}.")
        return

    has_errors = False
    for filename in os.listdir(input_dir):
        if filename.endswith('.pdf'):
            source_file = os.path.join(input_dir, filename)
            output_file = os.path.join(output_dir, filename)
            if convert_to_pdfa(source_file, output_file):
                os.remove(source_file)
            else:
                error_destination = os.path.join(error_dir, filename)
                shutil.move(source_file, error_destination)
                has_errors = True
                logging.warning(f"File moved to error directory: {filename}")

    if has_errors:
        logging.info("There has been at least one error, please check the PDF_Not_Converted folder.")


if __name__ == '__main__':
    base_path = get_base_path()
    input_directory = os.path.join(base_path, 'PDFA_IN')
    output_directory = os.path.join(base_path, 'PDFA_OUT')
    error_directory = os.path.join(base_path, 'PDF_Not_Converted')

    os.makedirs(output_directory, exist_ok=True)
    os.makedirs(error_directory, exist_ok=True)

    logging.info("Starting batch conversion process.")
    batch_convert(input_directory, output_directory, error_directory)
    logging.info("Batch conversion process completed.")
