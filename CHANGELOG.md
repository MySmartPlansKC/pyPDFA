# Changelog
All notable changes to this project will be documented in this file.

## [1.4.3] - 2024-06-07
### Changed
  - Adjusted Metadata creation to be more complete.

---

## [1.4.2] - 2024-06-05
### Fixed
  - Corrected error that was deleting the input directory "PDFA_IN" when processing files not in folders.

---

## [1.4.1] - 2024-06-05
### Changed
  - Adjusted logging to be more comprehensive with new structure.
  - Introduced a stack trace log for better tracking and to keep separate from general error logging.

---

## [1.4.0] - 2024-06-05
### Changed
  - Converted to pikepdf and Ghostscript pdf libraries for more comprehensive conversions.
  - Removed concurrent processing since it was causing issues at this time.

---

## [1.3.1] - 2024-05-21
### Fixed
  - bug causing the PDFA_IN folder to get deleted when removing the processed files.

---

## [1.3.0] - 2024-05-20
### Added
  - Recursive functionality so we can process folders inside the PDFA_IN folder

---

## [1.2.0] - 2024-05-14
### Added
  - New Package
    - Added pdfrw for new page count
    - Smaller package and better performance
    - 
### Removed
  - PymuPdf removed to reduce size and enhance performance

---

## [1.1.0] - 2024-05-09
### Added
- New feature
  - Page count 1/x display for better tracking

### Fixed
- Bug causing occasional crashes. 

---

## [1.0.0] - 2024-05-08
### Initial Release 