# CIE Certificate & UCAS Info Extractor

A Python application with a GUI for extracting student information from Cambridge International Examinations (CIE) statements and UCAS PDF documents, converting them into Excel format.

> This project was developed for my school's Student Guidance Office (SGO) to streamline their document processing workflow.

## Features

- **CIE Statement Extraction**: Parse CIE certificate/statement PDFs to extract:
  - Candidate name and number
  - School name and exam date
  - Subject names, grades, and syllabus codes
  - Percentage marks and qualification levels
  
- **UCAS Data Extraction**: Extract from UCAS PDF documents:
  - Student name and class group
  - Education history
  - School names, subjects, and grades
  - Qualification dates and awarding organizations

- **Dual PDF Support**: Handles both:
  - Electronic PDFs (text-based)
  - Scanned documents (using OCR)

- **Batch Processing**: Process multiple PDF files in a single operation

- **User-Friendly GUI**: Simple tkinter interface for selecting directories and monitoring progress

## Requirements

- Python 3.8+
- PyMuPDF (fitz)
- xlsxwriter
- more-itertools

## Installation

```bash
# Clone the repository
git clone https://github.com/anson/cie-certificate-and-ucas-info-exact.git
cd cie-certificate-and-ucas-info-exact

# Install dependencies
pip install pymupdf xlsxwriter more-itertools
```

## Usage

### Running the GUI

```bash
python main.py
```

1. Select the **CIE Statement Directory** containing CIE PDF files
2. Select the **UCAS PDF Directory** containing UCAS PDF files
3. Choose an **Output Directory** for the generated Excel files
4. Click **"Generate CIE Statement XLSX"** or **"Generate UCAS XLSX"** to process

### Running Individual Extractors

**CIE Extractor:**
```python
from parse_cie_statement import CambridgeOCRExtractor

extractor = CambridgeOCRExtractor(dpi=300)
records = extractor.extract("path/to/cie_statement.pdf")
extractor.write_to_xlsx(records, "cie_results.xlsx")
```

**UCAS Extractor:**
```python
from parse_ucas_statement import UCASExtractor

extractor = UCASExtractor("path/to/ucas.pdf")
data = extractor.extract()
extractor.write_to_xlsx(data, "ucas_results.xlsx")
```

## Building Executable

To build a standalone executable using PyInstaller:

```bash
pyinstaller main.spec
```

The executable will be created in the `dist/main/` directory.

## Output Format

### CIE Results (XLSX)

| Column | Description |
|--------|-------------|
| candidate_name | Student's full name |
| candidate_number | CIE candidate number |
| school | School/center name |
| exam_date | Examination date |
| document_type | Type of document |
| subject_name | Subject name |
| subject_grade | Grade achieved |
| subject_level | Qualification level |
| syllabus_code | Subject syllabus code |
| percentage_mark | Percentage mark (if available) |

### UCAS Results (XLSX)

| Column | Description |
|--------|-------------|
| name | Student's full name |
| group | Class/group identifier |
| school_name | Institution name |
| qualification_category | Type of qualification |
| subject_name | Subject name |
| subject_grade | Grade achieved |
| subject_date | Qualification date |
| subject_awarding_org | Awarding organization |
| subject_country | Country of awarding organization |

## Project Structure

```
.
├── main.py                    # Main GUI application
├── parse_cie_statement.py     # CIE PDF extraction logic
├── parse_ucas_statement.py    # UCAS PDF extraction logic
├── main.spec                  # PyInstaller configuration
└── README.md                  # This file
```

## License

MIT License

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.
