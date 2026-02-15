# PDF Extraction Script

A Python script that extracts invoice data from PDF files (Mogli Labs, SDI, JLL) and outputs structured data into an Excel file with separate worksheets per company type.

## Supported Company Types

| Company | Sheet Name | Description |
|---------|-----------|-------------|
| Mogli Labs (India) Private Limited | Mogli Lab | Product/goods invoices with unit pricing |
| SDI Business Services India Pvt Ltd | SDI | Product/goods invoices with rate-based pricing |
| Jones Lang LaSalle Property Consultants (India) Private Ltd | JLL | Service invoices with site and service type info |

## Features

- Automatically detects company type from PDF content
- Extracts structured invoice data (vendor, buyer, GST numbers, line items, tax breakdowns)
- Outputs Excel file with separate worksheets per company type
- Handles multi-page PDF invoices
- Processes all PDFs in a given input directory

## Installation

1. Clone this repository:
```bash
git clone https://github.com/sambit-star/pdf-extraction-script.git
cd pdf-extraction-script
```

2. Install the required dependencies:
```bash
pip install -r requirements.txt
```

## Usage

```bash
python pdf_extractor.py -i /path/to/input/folder -o /path/to/output
```

### Arguments

- `-i, --input-dir`: (Required) Path to the folder containing PDF invoice files
- `-o, --output-dir`: (Required) Directory where the Excel output will be saved. Will be created if it doesn't exist.

### Examples

Process PDFs in the current directory:
```bash
python pdf_extractor.py -i ./input -o ./output
```

Process PDFs with absolute paths:
```bash
python pdf_extractor.py -i /home/user/invoices -o /home/user/extracted_data
```

## Output Format

The script generates an Excel file (`extracted_invoices.xlsx`) with separate worksheets for each company type detected. Each worksheet has company-specific columns:

### Mogli Lab Sheet Columns
Invoice Type, Vendor Name (Billed From), Vendor GSTN, Buyer Name (Billed To), Buyer GSTIN, Place of Supply, Invoice No, Invoice Date, Description of Goods, HSN, Qty, Unit, Unit Price, Taxable Total, IGST %, IGST Amount, CGST %, CGST Amount, SGST %, SGST Amount, Amount

### SDI Sheet Columns
Invoice Type, Vendor Name (Billed From), Vendor GSTN, Buyer Name (Billed To), Buyer GSTIN, Invoice No, Invoice Date, Description of Goods, HSN/SAC, Quantity, Rate, Per, Taxable Total, IGST %, IGST Amount, CGST %, CGST Amount, SGST %, SGST Amount, Amount

### JLL Sheet Columns
Invoice Type, Vendor Name (Billed From), Vendor GSTN, Buyer Name (Billed To), Buyer GSTIN, Place of Supply, Invoice No, Invoice Date, Description, HSN, Qty, Taxable Value, IGST %, IGST Amount, CGST %, CGST Amount, SGST %, SGST Amount, Amount, Site Name, Service Type

## Requirements

- Python 3.6 or higher
- PyPDF2 3.0.0 or higher
- openpyxl 3.1.0 or higher

## License

This project is open source and available under the MIT License.