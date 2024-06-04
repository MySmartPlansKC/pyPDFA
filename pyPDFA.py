import pikepdf
import logging
import os
import shutil
import subprocess
import sys
import time
from typing import Union

# Versioning
__version__ = "1.4.0"
# pyinstaller --onefile --name pyPDFA-V1.4.0 pyPDFA.py

# Global logger variable
logger = logging.getLogger(__name__)


def setup_logging():
    global logger
    logger.setLevel(logging.DEBUG)

    # Stream handler for command line output
    stream_handler = logging.StreamHandler()
    stream_handler.setLevel(logging.INFO)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
    stream_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)

    return logger


def enable_error_logging():
    global logger
    if not any(isinstance(handler, logging.FileHandler) for handler in logger.handlers):
        parent_directory = os.path.dirname(input_directory)
        log_file_path = os.path.join(parent_directory, "pdf_conversion_error.log")
        error_file_handler = logging.FileHandler(log_file_path)
        error_file_handler.setLevel(logging.ERROR)
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
        error_file_handler.setFormatter(formatter)
        logger.addHandler(error_file_handler)


def safe_remove(path):
    try:
        os.unlink(path)
    except PermissionError:
        logger.error(f"Could not delete {path} - it may be in use.")
    except Exception as e:
        logger.error(f"Error deleting file {path}: {e}")


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
                logger.error(f"Could not delete directory {dir_path}: {str(e)}")


def get_base_path():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))


def get_pdf_page_count(pdf_path):
    try:
        with pikepdf.open(pdf_path) as pdf:
            page_count = len(pdf.pages)
            return page_count
    except Exception as e:
        logger.error(f"Error getting page count for {pdf_path}: {e}")
        enable_error_logging()
        return 0


def remove_annotations_and_comments(pdf_path):
    try:
        with pikepdf.open(pdf_path, allow_overwriting_input=True) as pdf:
            for page in pdf.pages:
                if '/Annots' in page:
                    del page['/Annots']
            pdf.save(pdf_path)
            logger.info(f"Removed annotations and comments from {pdf_path}")
    except Exception as e:
        logger.error(f"Failed to remove annotations and comments from {pdf_path}: {str(e)}")
        enable_error_logging()


def get_timeout(file_size_kb):
    if file_size_kb < 400000:
        return 300
    elif file_size_kb < 1000000:
        return 600
    else:
        return 900


def set_pdfa_metadata(pdf_path):
    try:
        with pikepdf.open(pdf_path, allow_overwriting_input=True) as pdf:
            info = pdf.docinfo
            creation_date = info.get('/CreationDate')
            info['/Producer'] = "MySmartPlans.com"
            info['/pdfaid:Conformance'] = "B"
            info['/pdfaid:Part'] = "1"
            if creation_date:
                info['/CreationDate'] = creation_date
            pdf.save()
            logger.info(f"Set PDF/A metadata for {pdf_path}")
    except Exception as e:
        logger.error(f"Failed to set PDF/A metadata for {pdf_path}: {str(e)}")
        enable_error_logging()


def convert_to_pdfa(source_path, output_path, error_dir, input_dir):
    process = None
    try:
        # Remove annotations and comments
        remove_annotations_and_comments(source_path)

        gs_executable = r'C:\Program Files\gs\gs10.03.0\bin\gswin64c.exe'
        page_count = get_pdf_page_count(source_path)
        file_size_kb = os.path.getsize(source_path) / 1024
        timeout_seconds = get_timeout(file_size_kb)

        cmd = [
            gs_executable,
            "-dQUIET",
            "-dPDFA=1",
            "-dBATCH",
            "-dNOPAUSE",
            "-dNOOUTERSAVE",
            "-sColorConversionStrategy=sRGB",
            "-sDEVICE=pdfwrite",
            "-dPDFACompatibilityPolicy=1",
            "-dEmbedAllFonts=true",
            "-dSubsetFonts=true",
            "-dConvertCMYKImagesToRGB=true",
            "-dRemoveAnnots=true",
            "-dRemoveComments=true",
            "-dWriteXRefStm=false",
            f"-sOutputFile={output_path}",
            source_path
        ]

        logger.info(f"Pages to convert: {page_count}")
        logger.info(f"Timeout set to {timeout_seconds} seconds for file size {file_size_kb:.0f} KB")

        process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        _, stderr = process.communicate(timeout=timeout_seconds)

        if process.returncode == 0:
            set_pdfa_metadata(output_path)
            logger.info(f"Successfully converted {source_path} to PDF/A.")
            return True, page_count
        else:
            logger.error(f"Ghostscript error: {stderr.decode()}")
            enable_error_logging()
            return False, page_count

    except subprocess.TimeoutExpired:
        logger.error(f"Timeout expired during conversion of {source_path}")
        enable_error_logging()
        return False, 0
    except subprocess.CalledProcessError as e:
        logger.error(f"Failed to convert {source_path}: {str(e)}")
        enable_error_logging()
        return False, 0
    finally:
        if process is not None:
            process.terminate()
            process.wait(timeout=5)
        if process.returncode != 0 or (process is not None and process.poll() is None):
            if os.path.exists(output_path):
                os.remove(output_path)
                logger.info(f"Rolled back changes: {output_path} removed")
            move_to_error_directory(source_path, error_dir, input_dir)


def move_to_error_directory(source_path, error_dir, input_dir):
    if not os.path.exists(error_dir):
        os.makedirs(error_dir)
    relative_path = os.path.relpath(str(source_path), input_dir)
    destination_path = os.path.join(error_dir, relative_path)

    os.makedirs(os.path.dirname(destination_path), exist_ok=True)
    shutil.move(str(source_path), str(destination_path))

    global logger
    logger.error(f"File moved to error directory: {relative_path}")
    enable_error_logging()


def check_and_clear_directory(directory):
    if os.path.exists(directory):
        if os.listdir(directory):  # Check if the directory is not empty
            response = input(
                f"Directory {directory} is not empty. Delete all contents? (y/n): "
            )
            if response.lower() == 'y':
                safe_rmtree(directory)
                logger.info(f"All contents of {directory} have been deleted.")
                time.sleep(1)
            else:
                logger.error("Operation aborted by the user.")
                enable_error_logging()
                return False
    return True


def remove_empty_directories(path, root_dir):
    try:
        if os.path.isdir(path) and not os.listdir(path):
            os.rmdir(path)
            logger.info(f"Removed empty directory: {path}")
            parent_dir = os.path.dirname(path)
            if parent_dir != root_dir:  # Ensure the root directory is not removed
                remove_empty_directories(parent_dir, root_dir)
    except OSError as e:
        logger.error(f"Error removing directory {path}: {str(e)}")
        enable_error_logging()


def batch_convert(input_dir: Union[str, os.PathLike], output_dir: Union[str, os.PathLike],
                  error_dir: Union[str, os.PathLike]):
    # Check and possibly clear the output directory
    if not check_and_clear_directory(output_dir):
        return
    time.sleep(1)

    # Check and possibly clear the error directory
    if not check_and_clear_directory(error_dir):
        return
    time.sleep(1)

    if not os.path.exists(input_dir) or not os.listdir(input_dir):
        logger.error(f"No files to process in {input_dir}.")
        enable_error_logging()
        return

    has_errors = False
    for root, dirs, files in os.walk(input_dir):
        for file in files:
            if file.endswith('.pdf'):
                source_file = os.path.join(root, file)
                relative_path = os.path.relpath(root, input_dir)
                output_file_dir = os.path.join(output_dir, relative_path)
                output_file = os.path.join(output_file_dir, file)

                os.makedirs(output_file_dir, exist_ok=True)
                success, _ = convert_to_pdfa(source_file, output_file, error_dir, input_dir)

                if not success:
                    if not has_errors:
                        enable_error_logging()
                        has_errors = True
                    error_file_dir = os.path.join(error_dir, relative_path)
                    os.makedirs(error_file_dir, exist_ok=True)
                    shutil.move(source_file, os.path.join(error_file_dir, file))
                    logger.warning(f"File moved to error directory: {file}")

                else:
                    os.remove(source_file)
                    remove_empty_directories(os.path.dirname(source_file), input_dir)

    if has_errors:
        logger.info("There has been at least one error, please check the PDF_Not_Converted folder.")


if __name__ == '__main__':
    base_path = get_base_path()

    # Testing paths
    # input_directory = os.path.join(base_path, "..", "xPDFTestFiles", "PDFA_IN")
    # output_directory = os.path.join(base_path, "..", "xPDFTestFiles", "PDFA_OUT")
    # error_directory = os.path.join(base_path, "..", "xPDFTestFiles", "PDF_Not_Converted")
    # input_directory = r"E:\Python\xPDFTestFiles\PDFA_IN"
    # output_directory = r"E:\Python\xPDFTestFiles\PDFA_OUT"
    # error_directory = r"E:\Python\xPDFTestFiles\PDF_Not_Converted"

    # Uncomment below for production paths
    input_directory = os.path.join(base_path, "PDFA_IN")
    output_directory = os.path.join(base_path, "PDFA_OUT")
    error_directory = os.path.join(base_path, "PDF_Not_Converted")

    logger = setup_logging()

    logger.info(f"Starting PDFA Conversion v{__version__}")
    time.sleep(1)

    os.makedirs(output_directory, exist_ok=True)

    batch_convert(input_directory, output_directory, error_directory)
    logger.info("Batch conversion process completed.")
