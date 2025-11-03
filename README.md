# pdf-to-excel-automation
Developed a Python automation script at Pinchin Ltd. that extracted data from PDFs and exported it to Excel, significantly reducing manual data entry time.

# PDF to Excel Automation

This project automates the workflow of taking structured data out of PDF field or environmental reports and exporting it into a clean Excel template for analysis and storage.

It was originally built to reduce manual data entry time for laboratory staff by parsing common report formats and writing the results into spreadsheets.

## Features
- Extracts tabular or semi-structured data from PDFs
- Maps extracted fields to a defined Excel layout
- Skips duplicate or malformed entries to protect data integrity
- Logs processing results (success, skipped, errors)
- Designed to be extended for new PDF templates

## Tech Stack
- Python 3.x
- PDF parsing: `pdfplumber` 
- Excel writing: `openpyxl`

## Project Structure
```text
src/
  main.py          # CLI / entry point
  pdf_parser.py    # PDF extraction logic
  excel_writer.py  # Excel export logic
  utils.py         # helpers (logging, validation)
samples/
  sample_report.pdf
  output_example.xlsx
