#!/usr/bin/env python3
"""
PDF Extraction Script
Extracts data from PDF files and outputs JSON representation
"""

import json
import os
import sys
from pathlib import Path
import argparse

try:
    import PyPDF2
except ImportError:
    print("Error: PyPDF2 is not installed. Please install it using: pip install PyPDF2")
    sys.exit(1)


def extract_pdf_data(pdf_path):
    """
    Extract data from a PDF file.
    
    Args:
        pdf_path: Path to the PDF file
        
    Returns:
        Dictionary containing extracted data
    """
    extracted_data = {
        "file_name": os.path.basename(pdf_path),
        "file_path": str(pdf_path),
        "pages": [],
        "metadata": {},
        "total_pages": 0
    }
    
    try:
        with open(pdf_path, 'rb') as pdf_file:
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            
            # Extract metadata
            if pdf_reader.metadata:
                metadata = {}
                for key, value in pdf_reader.metadata.items():
                    # Convert metadata key to string and remove leading '/'
                    clean_key = str(key).lstrip('/')
                    metadata[clean_key] = str(value) if value else None
                extracted_data["metadata"] = metadata
            
            # Extract total pages
            extracted_data["total_pages"] = len(pdf_reader.pages)
            
            # Extract text from each page
            for page_num, page in enumerate(pdf_reader.pages, start=1):
                page_text = page.extract_text()
                page_data = {
                    "page_number": page_num,
                    "text": page_text.strip() if page_text else ""
                }
                extracted_data["pages"].append(page_data)
        
        return extracted_data
    
    except FileNotFoundError:
        print(f"Error: PDF file not found at {pdf_path}")
        sys.exit(1)
    except PermissionError:
        print(f"Error: Permission denied when trying to read {pdf_path}")
        sys.exit(1)
    except PyPDF2.errors.PdfReadError as e:
        print(f"Error: Invalid or corrupted PDF file - {str(e)}")
        sys.exit(1)
    except Exception as e:
        print(f"Error: Unexpected error while extracting data from PDF - {str(e)}")
        sys.exit(1)


def save_json(data, output_dir, pdf_filename):
    """
    Save extracted data as JSON file.
    
    Args:
        data: Dictionary containing extracted data
        output_dir: Directory to save the JSON file
        pdf_filename: Original PDF filename (used to create JSON filename)
    """
    # Create output filename (replace .pdf with .json)
    json_filename = os.path.splitext(pdf_filename)[0] + ".json"
    output_path = os.path.join(output_dir, json_filename)
    
    try:
        with open(output_path, 'w', encoding='utf-8') as json_file:
            json.dump(data, json_file, indent=2, ensure_ascii=False)
        
        print(f"✓ Successfully extracted data from PDF")
        print(f"✓ JSON output saved to: {output_path}")
    
    except PermissionError:
        print(f"Error: Permission denied when trying to write to {output_path}")
        sys.exit(1)
    except OSError as e:
        print(f"Error: Failed to write JSON file - {str(e)}")
        sys.exit(1)
    except Exception as e:
        print(f"Error: Unexpected error while saving JSON file - {str(e)}")
        sys.exit(1)


def main():
    """Main function to orchestrate PDF extraction."""
    parser = argparse.ArgumentParser(
        description='Extract data from PDF files and output as JSON',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python pdf_extractor.py --pdf-path /path/to/file.pdf --output-dir /path/to/output
  python pdf_extractor.py -p document.pdf -o ./output
        """
    )
    
    parser.add_argument(
        '-p', '--pdf-path',
        required=True,
        help='Path to the PDF file to extract data from'
    )
    
    parser.add_argument(
        '-o', '--output-dir',
        required=True,
        help='Directory to save the JSON output (will be created if it does not exist)'
    )
    
    args = parser.parse_args()
    
    # Validate PDF file exists
    pdf_path = Path(args.pdf_path)
    if not pdf_path.exists():
        print(f"Error: PDF file not found: {args.pdf_path}")
        sys.exit(1)
    
    if not pdf_path.is_file():
        print(f"Error: Path is not a file: {args.pdf_path}")
        sys.exit(1)
    
    if pdf_path.suffix.lower() != '.pdf':
        print(f"Error: File is not a PDF: {args.pdf_path}")
        sys.exit(1)
    
    # Create output directory if it doesn't exist
    output_dir = Path(args.output_dir)
    try:
        output_dir.mkdir(parents=True, exist_ok=True)
        print(f"✓ Output directory ready: {output_dir}")
    except Exception as e:
        print(f"Error creating output directory: {str(e)}")
        sys.exit(1)
    
    # Extract data from PDF
    print(f"Extracting data from: {pdf_path}")
    extracted_data = extract_pdf_data(pdf_path)
    
    # Save as JSON
    save_json(extracted_data, str(output_dir), pdf_path.name)


if __name__ == "__main__":
    main()
