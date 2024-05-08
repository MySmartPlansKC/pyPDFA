# PDFA Conversion Tool

## Overview
This tool automates the conversion of PDF files to PDF/A format to ensure they meet archival standards.

## Version
1.0.0

## Setup
Before running the executable, make sure the following directories are present:
- `PDFA_IN`: Place your PDFs here for conversion.
- `PDFA_OUT`: Converted PDF/A files will be saved here.
- `PDF_Not_Converted`: PDFs that failed conversion will be moved here.

The application checks and creates these directories if they don't exist.

## Usage
1. **Prepare Your Files**:
   - Place PDF files into the `PDFA_IN` directory.
2. **Run the Tool**:
   - Execute the tool. It processes all files from `PDFA_IN`, attempts to convert them, and handles them based on success or failure.
3. **Check for Errors**:
   - Check the `PDF_Not_Converted` directory for any files that could not be converted and review the `pdf_conversion.log` for detailed logs.
4. **Review Outputs**:
   - Converted files can be found in the `PDFA_OUT` directory.

## Dependencies
- **Ghostscript 10.03.0** must be installed on your machine. Ensure the executable (`gswin64c.exe`) is properly referenced in the script.

## License
This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details.

## Support
For support, please open an issue in the GitHub issue tracker.
