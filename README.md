# PDF Extraction Script

A Python script that extracts data from PDF files and outputs it in JSON format.

## Features

- Extracts text content from all pages of a PDF
- Extracts PDF metadata (author, title, creation date, etc.)
- Outputs data in well-structured JSON format
- Automatically creates output directory if it doesn't exist
- Command-line interface for easy usage

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

Basic usage:
```bash
python pdf_extractor.py --pdf-path /path/to/your/file.pdf --output-dir /path/to/output
```

Or using short options:
```bash
python pdf_extractor.py -p document.pdf -o ./output
```

### Arguments

- `-p, --pdf-path`: (Required) Path to the PDF file you want to extract data from
- `-o, --output-dir`: (Required) Directory where the JSON output will be saved. Will be created if it doesn't exist.

### Examples

Extract data from a PDF in the current directory:
```bash
python pdf_extractor.py -p mydocument.pdf -o ./output
```

Extract data from a PDF with absolute path:
```bash
python pdf_extractor.py -p /home/user/documents/report.pdf -o /home/user/extracted_data
```

## Output Format

The script generates a JSON file with the following structure:

```json
{
  "file_name": "document.pdf",
  "file_path": "/path/to/document.pdf",
  "total_pages": 5,
  "metadata": {
    "Title": "Sample Document",
    "Author": "John Doe",
    "CreationDate": "D:20240101120000",
    "Producer": "PDF Generator"
  },
  "pages": [
    {
      "page_number": 1,
      "text": "Extracted text from page 1..."
    },
    {
      "page_number": 2,
      "text": "Extracted text from page 2..."
    }
  ]
}
```

The JSON file will have the same name as the input PDF file (with .json extension) and will be saved in the specified output directory.

## Requirements

- Python 3.6 or higher
- PyPDF2 3.0.0 or higher

## License

This project is open source and available under the MIT License.