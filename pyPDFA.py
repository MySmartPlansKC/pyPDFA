import logging
import os
import pikepdf
import shutil
import subprocess
import sys
import time
from pathlib import Path
from typing import Union
from pikepdf.models.metadata import PdfMetadata

# Versioning
__version__ = "1.5.0"
# pyinstaller --onefile --name pyPDFA-V1.5.0 pyPDFA.py

# Global variables
logger = logging.getLogger('main_logger')
error_logger = logging.getLogger('error_logger')
stacktrace_logger = logging.getLogger('stacktrace_logger')
gs_logger = logging.getLogger('ghostscript_logger')
gs_log_file_path = None
base_path = None


def setup_logging():
    global gs_log_file_path
    logger.setLevel(logging.DEBUG)

    # Console Handler
    stream_handler = logging.StreamHandler()
    stream_handler.setLevel(logging.INFO)
    formatter = logging.Formatter(
        '%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    stream_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)

    # File Handler for Errors
    error_log_file_path = base_path / "pdf_conversion_error.log"
    error_file_handler = logging.FileHandler(error_log_file_path)
    error_file_handler.setLevel(logging.ERROR)
    error_file_handler.setFormatter(formatter)
    error_logger.addHandler(error_file_handler)

    # File Handler for Stack Traces
    stacktrace_log_file_path = base_path / "pdf_conversion_stacktrace.log"
    stacktrace_file_handler = logging.FileHandler(stacktrace_log_file_path)
    stacktrace_file_handler.setLevel(logging.ERROR)
    stacktrace_file_handler.setFormatter(formatter)
    stacktrace_logger.addHandler(stacktrace_file_handler)

    # File Handler for Ghostscript Logs
    gs_log_file_path = base_path / "ghostscript.log"
    gs_file_handler = logging.FileHandler(gs_log_file_path)
    gs_file_handler.setLevel(logging.DEBUG)  # Capture all Ghostscript logs
    gs_file_handler.setFormatter(formatter)
    gs_logger.addHandler(gs_file_handler)
    gs_logger.setLevel(logging.DEBUG)


def log_error(message: str):
    error_logger.error(message)


def log_exception(message: str):
    stacktrace_logger.exception(message)


def safe_remove(path: Path):
    try:
        path.unlink()
    except PermissionError:
        log_error(f"Could not delete {path} - it may be in use.")
        log_exception(f"Error deleting file {path}:")
    except Exception as e:
        log_error(f"Error deleting file {path}: {e}")
        log_exception("Stack trace:")


def clear_input_directory(directory: Path):
    for root, dirs, files in os.walk(directory, topdown=False):
        root_path = Path(root)
        for name in files:
            file_path = root_path / name
            safe_remove(file_path)
        for name in dirs:
            dir_path = root_path / name
            if "PDFA_IN" not in dir_path.name:
                try:
                    dir_path.rmdir()
                except OSError as e:
                    log_error(f"Could not delete directory {dir_path}: {e}")
                    log_exception("Stack trace:")


def safe_rmtree(directory: Path):
    for root, dirs, files in os.walk(directory, topdown=False):
        root_path = Path(root)
        for name in files:
            file_path = root_path / name
            safe_remove(file_path)
        for name in dirs:
            dir_path = root_path / name
            try:
                dir_path.rmdir()
            except OSError as e:
                log_error(f"Could not delete directory {dir_path}: {e}")
                log_exception("Stack trace:")


def get_base_path(testing_dir: Union[Path, None] = None) -> Path:
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    elif testing_dir:
        return Path(testing_dir)
    else:
        return Path.cwd()


def get_pdf_page_count(pdf_path: Path) -> int:
    try:
        with pikepdf.open(pdf_path) as pdf:
            return len(pdf.pages)
    except Exception as e:
        log_error(f"Error getting page count for {pdf_path}: {e}")
        log_exception("Stack trace:")
        return 0


def remove_annotations_and_comments(pdf_path: Path):
    try:
        with pikepdf.open(pdf_path, allow_overwriting_input=True) as pdf:
            for page in pdf.pages:
                if '/Annots' in page:
                    del page['/Annots']
            pdf.save(pdf_path)
    except Exception as e:
        log_error(f"Failed to remove annotations and comments from {pdf_path}: {e}")
        log_exception("Stack trace:")


def get_timeout(file_size_kb: float) -> int:
    if file_size_kb < 150000:
        return 480  # 8 minutes
    elif file_size_kb < 300000:
        return 900  # 15 minutes
    elif file_size_kb < 600000:
        return 2400  # 40 minutes
    elif file_size_kb < 900000:
        return 3600  # 60 minutes
    else:
        return 5400  # 90 minutes


def set_pdfa_metadata(pdf_path: Path):
    try:
        with pikepdf.open(pdf_path, allow_overwriting_input=True) as pdf:
            with pdf.open_metadata(set_pikepdf_as_editor=False, update_docinfo=True, strict=True) as meta:
                meta['xmp:CreatorTool'] = f"pyPDFA v{__version__}"
                meta['pdf:Producer'] = "MySmartPlans.com"
                meta['pdfaid:Conformance'] = "B"
                meta['pdfaid:Part'] = "1"  # PDF/A-1b Compliance

                if 'dc:title' in meta and meta['dc:title'] == 'Untitled':
                    del meta['dc:title']

            _unset_empty_metadata(meta)
            pdf.save()
    except Exception as e:
        log_error(f"Failed to set PDF/A metadata for {pdf_path}: {e}")
        log_exception("Stack trace:")


def _unset_empty_metadata(meta: PdfMetadata):
    """Unset metadata fields that were explicitly set to empty strings."""
    fields = ['dc:title', 'dc:creator', 'pdf:Author', 'dc:description', 'dc:subject', 'pdf:Keywords']
    for field in fields:
        if field in meta and not meta[field]:
            del meta[field]


def convert_to_pdfa(
        source_path: Path,
        output_path: Path,
        error_dir: Path,
        input_dir: Path,
        document_index: int,
        total_documents: int
) -> bool:
    process = None
    page_count = get_pdf_page_count(source_path)
    file_size_kb = source_path.stat().st_size / 1024
    timeout_seconds = get_timeout(file_size_kb)
    timeout_minutes = timeout_seconds / 60

    try:
        # Remove annotations and comments
        remove_annotations_and_comments(source_path)

        gs_executable = r'C:\Program Files\gs\gs10.03.0\bin\gswin64c.exe'
        # icc_profile_path = icc_profile_directory / icc_profile

        # if not icc_profile_path.is_file():
        #     log_error(f"ICC profile not found at {icc_profile_path}. Conversion aborted.")
        #     logger.error(f"ICC profile not found at {icc_profile_path}.")
        #     return False

        cmd = [
            gs_executable,
            "-dPDFA=1",
            "-dQUIET",
            "-dNOOUTERSAVE",
            "-dBATCH",
            "-dNOPAUSE",
            "-dPDFACompatibilityPolicy=1",
            "-dEmbedAllFonts=true",
            "-dSubsetFonts=true",
            "-dCompressFonts=true",
            "-dAutoFilterColorImages=true",
            "-dAutoFilterGrayImages=true",
            "-dColorImageFilter=/FlateEncode",
            "-dGrayImageFilter=/FlateEncode",
            "-dMonoImageFilter=/CCITTFaxEncode",
            "-dCompressPages=true",
            "-dDownsampleColorImages=true",
            "-dDownsampleGrayImages=true",
            "-dDownsampleMonoImages=true",
            # "-dColorImageResolution=150",
            # "-dGrayImageResolution=150",
            # "-dMonoImageResolution=150",
            # "-dFastWebView=true",
            "-sDEVICE=pdfwrite",
            "-sColorConversionStrategy=RGB",
            # f"-sOutputICCProfile={str(icc_profile_path)}",
            f"-sOutputFile={str(output_path)}",
            str(source_path)
        ]

        # Log the command for debugging
        # logger.info(f"Running Ghostscript command: {' '.join(cmd)}")
        logger.info(f"Processing document {document_index} of {total_documents}")
        logger.info(f"Timeout set to {timeout_minutes} minutes for file size {file_size_kb:.0f} KB")
        logger.info(f"Pages to convert: {page_count}")
        logger.info(f"Please wait for the conversion to complete...")

        # Redirect stdout & stderr to the Ghostscript log file
        with open(str(gs_log_file_path), "a") as gs_log:
            process = subprocess.Popen(
                cmd,
                stdout=gs_log,
                stderr=gs_log,
                text=True
            )
            process.communicate(timeout=timeout_seconds)

        # After Ghostscript finishes, check if the output file exists
        if process.returncode == 0:
            set_pdfa_metadata(output_path)  # Set metadata after successful conversion
            logger.info(f"Successfully converted {source_path} to PDF/A.")
            return True
        else:
            # If exit code is non-zero, check if output file exists
            if output_path.exists():
                set_pdfa_metadata(output_path)
                logger.warning(f"""Conversion for
                    {source_path}
                    completed with warnings.""")
                log_error(
                    f"Ghostscript returned exit code {process.returncode} for {source_path}, but output file exists.")
                return True
            else:
                # Treat as failure if output file does not exist
                log_error(f"Conversion failed for {source_path}; output file not found.")
                stacktrace_logger.error(f"Conversion failed for {source_path}: Exit code {process.returncode}")
                move_to_error_directory(source_path, error_dir, input_dir, output_path)
                return False

    except subprocess.TimeoutExpired:
        if process:
            process.terminate()
            process.wait()
        if output_path.exists():
            output_path.unlink()
        move_to_error_directory(source_path, error_dir, input_dir, output_path)
        logger.error(f"Timeout expired during conversion of {source_path}")
        log_error(f"Timeout expired during conversion of {source_path}")
        stacktrace_logger.error(f"Timeout expired during conversion of {source_path}")
        stacktrace_logger.error(f"{timeout_minutes} minute Timeout for file size {file_size_kb:.0f} KB")
        return False

    except subprocess.CalledProcessError as e:
        if process and process.poll() is None:
            process.terminate()
            process.wait()
        move_to_error_directory(source_path, error_dir, input_dir, output_path)
        log_error(f"Conversion error for {source_path}: {e}")
        stacktrace_logger.error(f"Conversion error for {source_path}: {e}")
        return False

    except Exception as e:
        if process and process.poll() is None:
            process.terminate()
            process.wait()
        log_error(f"Unexpected error during conversion of {source_path}: {e}")
        stacktrace_logger.error(f"Unexpected error during conversion of {source_path}: {e}")
        move_to_error_directory(source_path, error_dir, input_dir, output_path)
        return False

    finally:
        if process and process.poll() is None:
            process.terminate()
            process.wait()


def move_to_error_directory(source_path: Path, error_dir: Path, input_dir: Path, output_path: Path):
    time.sleep(1)
    try:
        if output_path.exists():
            output_path.unlink()
        relative_path = source_path.relative_to(input_dir)
        destination_path = error_dir / relative_path
        destination_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(source_path), str(destination_path))
        logger.info(f"File moved to error directory: {destination_path}")
        log_error(f"File moved to error directory: {destination_path}")
    except Exception as e:
        log_error(f"Failed to move file to error directory: {str(e)}")
        log_exception("Stack trace:")


def check_and_clear_directory(directory: Path) -> bool:
    if directory.exists():
        if any(directory.iterdir()):  # Check if the directory is not empty
            while True:
                response = input(f"Directory {directory} is not empty. Delete all contents? (Y/n): ").strip().lower()
                if response in ('y', 'yse', ''):
                    safe_rmtree(directory)
                    logger.info(f"All contents of {directory} have been deleted.")
                    break
                elif response in ('n', 'no'):
                    logger.error("Operation aborted by the user.")
                    return False
                else:
                    print("Invalid input. Please enter 'Y' for Yes or 'N' for No.")
    return True


def remove_empty_directories(path: Path, root_dir: Path):
    try:
        if path.is_dir() and not any(path.iterdir()):
            if path.name != "PDFA_IN":
                path.rmdir()
                parent_dir = path.parent
                if parent_dir != root_dir:
                    remove_empty_directories(parent_dir, root_dir)
    except OSError as e:
        log_error(f"Error removing directory {path}: {e}")
        log_exception("Stack trace:")


def batch_convert(input_dir: Path, output_dir: Path, error_dir: Path):
    if not check_and_clear_directory(output_dir):
        return
    if not check_and_clear_directory(error_dir):
        return
    if not input_dir.exists() or not any(input_dir.iterdir()):
        logger.error(f"No files to process in {input_dir}.")
        return

    total_documents = sum(1 for _ in input_dir.rglob('*.pdf'))
    document_index = 0

    has_errors = False
    for source_file in input_dir.rglob('*.pdf'):
        document_index += 1
        relative_path = source_file.relative_to(input_dir)
        output_file = output_dir / relative_path

        output_file.parent.mkdir(parents=True, exist_ok=True)
        success = convert_to_pdfa(
            source_file,
            output_file,
            error_dir,
            input_dir,
            document_index,
            total_documents
        )
        if not success:
            has_errors = True
        else:
            source_file.unlink()

        remove_empty_directories(source_file.parent, input_dir)

    clear_input_directory(input_dir)
    logger.info("Batch conversion process completed.")
    if has_errors:
        logger.info("There has been at least one error, please check the PDF_Not_Converted folder.")
        input("Press Enter to continue...")


if __name__ == '__main__':
    # Configuration flag to switch environments
    USE_PRODUCTION_PATHS = True  # Set to True for production & False for testing

    if USE_PRODUCTION_PATHS:
        # Production paths
        base_path = get_base_path()
        input_directory = base_path / "PDFA_IN"
        output_directory = base_path / "PDFA_OUT"
        error_directory = base_path / "PDF_Not_Converted"
    else:
        # Testing paths
        testing_directory = Path(r"E:\Python\xPDFTestFiles")
        base_path = get_base_path(testing_directory)
        input_directory = base_path / "PDFA_IN"
        output_directory = base_path / "PDFA_OUT"
        error_directory = base_path / "PDF_Not_Converted"

    setup_logging()

    logger.info(f"Starting PDFA Conversion v{__version__}")
    time.sleep(1)

    # Ensure directories exist
    input_directory.mkdir(parents=True, exist_ok=True)
    output_directory.mkdir(parents=True, exist_ok=True)

    batch_convert(input_directory, output_directory, error_directory)
