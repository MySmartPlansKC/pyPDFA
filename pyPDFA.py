import pikepdf
import logging
import os
import shutil
import subprocess
import sys
import time
from typing import Union

# Versioning
__version__ = "1.4.1"
# pyinstaller --onefile --name pyPDFA-V1.4.1 pyPDFA.py

# Global logger variables
logger = logging.getLogger('main_logger')
error_logger = logging.getLogger('error_logger')
stacktrace_logger = logging.getLogger('stacktrace_logger')


def setup_logging():
    logger.setLevel(logging.DEBUG)
    stream_handler = logging.StreamHandler()
    stream_handler.setLevel(logging.INFO)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
    stream_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)

    # Initialize placeholders for file handlers which will be setup on first error
    error_logger.setLevel(logging.ERROR)
    stacktrace_logger.setLevel(logging.ERROR)


def setup_file_logging(logging_base_path):
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')

    # Setup file handler for errors if not already set
    if not any(isinstance(h, logging.FileHandler) for h in error_logger.handlers):
        error_log_file_path = os.path.join(logging_base_path, "pdf_conversion_error.log")
        error_file_handler = logging.FileHandler(error_log_file_path)
        error_file_handler.setLevel(logging.ERROR)
        error_file_handler.setFormatter(formatter)
        error_logger.addHandler(error_file_handler)
        error_logger.error("Error log file setup complete.")

    # Setup file handler for stack traces if not already set
    if not any(isinstance(h, logging.FileHandler) for h in stacktrace_logger.handlers):
        stacktrace_log_file_path = os.path.join(logging_base_path, "pdf_conversion_stacktrace.log")
        stacktrace_file_handler = logging.FileHandler(stacktrace_log_file_path)
        stacktrace_file_handler.setLevel(logging.ERROR)
        stacktrace_file_handler.setFormatter(formatter)
        stacktrace_logger.addHandler(stacktrace_file_handler)
        stacktrace_logger.error("Stack trace log file setup complete.")


def log_error(message):
    setup_file_logging(base_path)
    error_logger.error(message)


def log_exception(message):
    setup_file_logging(base_path)
    stacktrace_logger.exception(message)


def safe_remove(path):
    try:
        os.unlink(path)
    except PermissionError:
        log_error(f"Could not delete {path} - it may be in use.")
        log_exception(f"Error deleting file {path}:")
    except Exception as e:
        log_error(f"Error deleting file {path}: {e}")
        log_exception("Stack trace:")


def clear_input_directory(directory):
    for root, dirs, files in os.walk(directory, topdown=False):
        for name in files:
            file_path = os.path.join(root, name)
            safe_remove(file_path)
        for name in dirs:
            dir_path = os.path.join(root, name)
            if "PDFA_IN" not in dir_path:
                try:
                    os.rmdir(dir_path)
                except OSError as e:
                    log_error(f"Could not delete directory {dir_path}: {e}")
                    log_exception("Stack trace:")


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
                log_error(f"Could not delete directory {dir_path}: {e}")
                log_exception("Stack trace:")


def get_base_path(input_directory_dev):
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(input_directory_dev)


def get_pdf_page_count(pdf_path):
    try:
        with pikepdf.open(pdf_path) as pdf:
            page_count = len(pdf.pages)
            return page_count
    except Exception as e:
        log_error(f"Error getting page count for {pdf_path}: {e}")
        log_exception("Stack trace:")
        return 0


def remove_annotations_and_comments(pdf_path):
    try:
        with pikepdf.open(pdf_path, allow_overwriting_input=True) as pdf:
            for page in pdf.pages:
                if '/Annots' in page:
                    del page['/Annots']
            pdf.save(pdf_path)
    except Exception as e:
        log_error(f"Failed to remove annotations and comments from {pdf_path}: {e}")
        log_exception("Stack trace:")


def get_timeout(file_size_kb):
    if file_size_kb < 150000:
        return 30  # 5 minutes
    elif file_size_kb < 300000:
        return 60  # 10 minutes
    elif file_size_kb < 600000:
        return 900  # 15 minutes
    elif file_size_kb < 900000:
        return 1200  # 20 minutes
    else:
        return 1500  # 25 minutes


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
            # logger.info(f"Set PDF/A metadata for {pdf_path}")
    except Exception as e:
        log_error(f"Failed to set PDF/A metadata for {pdf_path}: {e}")
        log_exception("Stack trace:")


def convert_to_pdfa(source_path, output_path, error_dir, input_dir):
    process = None
    page_count = get_pdf_page_count(source_path)
    file_size_kb = os.path.getsize(source_path) / 1024
    timeout_seconds = get_timeout(file_size_kb)
    try:
        # Remove annotations and comments
        remove_annotations_and_comments(source_path)

        gs_executable = r'C:\Program Files\gs\gs10.03.0\bin\gswin64c.exe'

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

        logger.info(f"Timeout set to {timeout_seconds} seconds for file size {file_size_kb:.0f} KB")
        logger.info(f"Pages to convert: {page_count}")
        logger.info(f"Please wait for the conversion to complete...")

        process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        _, stderr = process.communicate(timeout=timeout_seconds)

        if process.returncode == 0:
            set_pdfa_metadata(output_path)  # Set metadata after successful conversion
            logger.info(f"Successfully converted {source_path} to PDF/A.")
            return True
        else:
            logger.info(f"Conversion failed for {source_path}")
            log_error(f"Conversion failed for {source_path}")
            stacktrace_logger.error(f"Conversion failed for {source_path}: {stderr.decode()}")
            return False

    except subprocess.TimeoutExpired:
        if process:
            process.terminate()
            process.wait()
        if os.path.exists(output_path):
            os.remove(output_path)

        move_to_error_directory(source_path, error_dir, input_dir, output_path)
        logger.info(f"Timeout expired during conversion of {source_path}")
        log_error(f"Timeout expired during conversion of {source_path}")
        stacktrace_logger.error(f"Timeout expired during conversion of {source_path}")
        stacktrace_logger.error(f"{timeout_seconds} second Timeout for file size {file_size_kb:.0f} KB")
        return False

    except subprocess.CalledProcessError as e:
        if process and process.poll() is None:
            process.terminate()
            process.wait()

        move_to_error_directory(source_path, error_dir, input_dir, output_path)
        log_error(f"Conversion error for {source_path}: {e}")
        return False

    except Exception as e:
        if process and process.poll() is None:
            process.terminate()
            process.wait()

        log_error(f"Unexpected error during conversion of {source_path}: {e}")
        move_to_error_directory(source_path, error_dir, input_dir, output_path)
        return False

    finally:
        if process and process.poll() is None:
            process.terminate()
            process.wait()


def move_to_error_directory(source_path, error_dir, input_dir, output_path):
    time.sleep(1)
    try:
        if os.path.exists(output_path):
            os.remove(output_path)
        if not os.path.exists(error_dir):
            os.makedirs(error_dir)
        relative_path = os.path.relpath(str(source_path), input_dir)
        destination_path = os.path.join(error_dir, relative_path)

        os.makedirs(os.path.dirname(destination_path), exist_ok=True)
        shutil.move(str(source_path), str(destination_path))

        logger.info(f"File moved to error directory: {destination_path}")
        log_error(f"File moved to error directory: {destination_path}")
    except Exception as e:
        log_error(f"Failed to move file to error directory: {str(e)}")
        log_exception("Stack trace:")


def check_and_clear_directory(directory):
    if os.path.exists(directory):
        if os.listdir(directory):  # Check if the directory is not empty
            response = input(f"Directory {directory} is not empty. Delete all contents? (y/n): ")
            if response.lower() == 'y':
                safe_rmtree(directory)
                logger.info(f"All contents of {directory} have been deleted.")
            else:
                logger.error("Operation aborted by the user.")
                return False
    return True


def remove_empty_directories(path, root_dir):
    try:
        if os.path.isdir(path) and not os.listdir(path):
            if os.path.basename(path) != "PDFA_IN":
                os.rmdir(path)
                # logger.info(f"Removed empty directory: {path}")
                parent_dir = os.path.dirname(path)
                if parent_dir != root_dir:
                    remove_empty_directories(parent_dir, root_dir)
    except OSError as e:
        log_error(f"Error removing directory {path}: {e}")
        log_exception("Stack trace:")


def batch_convert(input_dir: Union[str, os.PathLike], output_dir: Union[str, os.PathLike], error_dir: Union[str, os.PathLike]):
    if not check_and_clear_directory(output_dir):
        return
    if not check_and_clear_directory(error_dir):
        return
    if not os.path.exists(input_dir) or not os.listdir(input_dir):
        logger.error(f"No files to process in {input_dir}.")
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
                success = convert_to_pdfa(source_file, output_file, error_dir, input_dir)
                if not success:
                    has_errors = True
                else:
                    os.remove(source_file)

                remove_empty_directories(os.path.dirname(source_file), input_dir)

    clear_input_directory(input_dir)

    logger.info("Batch conversion process completed.")
    if has_errors:
        logger.info("There has been at least one error, please check the PDF_Not_Converted folder.")
        input("Press Enter to continue...")


if __name__ == '__main__':
    # Testing paths
    input_directory = r"E:\Python\xPDFTestFiles\PDFA_IN"
    output_directory = r"E:\Python\xPDFTestFiles\PDFA_OUT"
    error_directory = r"E:\Python\xPDFTestFiles\PDF_Not_Converted"

    base_path = get_base_path(input_directory)
    print(f"The base path is: {base_path}")
    setup_logging()

    # Uncomment below for production paths
    # input_directory = os.path.join(base_path, "PDFA_IN")
    # output_directory = os.path.join(base_path, "PDFA_OUT")
    # error_directory = os.path.join(base_path, "PDF_Not_Converted")

    logger.info(f"Starting PDFA Conversion v{__version__}")
    time.sleep(1)

    if not os.path.exists(input_directory):
        os.makedirs(input_directory)
        print("The 'PDFA_IN' folder was missing and has now been created.")
        print("Please add your files/folders to the 'PDFA_IN' folder for processing.")
        input("Press Enter to close the program...")
        sys.exit(1)

    os.makedirs(output_directory, exist_ok=True)
    batch_convert(input_directory, output_directory, error_directory)
