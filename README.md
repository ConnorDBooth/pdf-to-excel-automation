# PDF to Excel Automation

Developed a Python-based automation tool at Pinchin Ltd. to extract semi-structured data from laboratory and environmental PDF reports and populate standardized Excel templates for analysis and record keeping.

The tool replaces manual transcription workflows by parsing common report formats, validating extracted data, and computing summary statistics, significantly reducing data entry time and error rates for laboratory staff.

Note: Real-world examples have been removed to maintain client and company confidentiality

## Features
- Extracts tabular and semi-structured data from PDF reports
- Automatically identifies relevant report sections (e.g. Outdoor samples)
- Maps extracted fields into a predefined Excel template
- Computes summary statistics (mean, standard deviation, percentiles, frequency)
- Handles missing, malformed, or duplicate entries to protect data integrity
- User-friendly GUI for non-technical users
- Designed to be extensible for new PDF templates

## Tech Stack
- Python 3.x
- PDF parsing: pdfplumber
- Excel automation: openpyxl
- GUI: tkinter
- Data validation and statistical analysis

## Project Structure
```text
src/
  main.py          # GUI and application entry point
  pdf_parser.py    # PDF table detection and extraction
  excel_writer.py  # Excel population and statistics computation
samples/
  Example.xlsx
```
## Design Challenges
- Handling inconsistent PDF table layouts across reports
- Distinguishing between zero values and missing data
- Dynamically inserting Excel columns without breaking existing formulas