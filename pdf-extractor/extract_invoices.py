"""
extract_invoices.py
-------------------
Batch PDF extractor — reads a folder of invoice/receipt PDFs
and outputs a clean CSV with extracted fields.

Usage:
    python extract_invoices.py --folder ./invoices --output results.csv

    Requirements:
        pip install pdfplumber pandas openpyxl
        """

import pdfplumber
import pandas as pd
import os
import argparse
import re


def extract_text_from_pdf(pdf_path):
      """Extract raw text from all pages of a PDF."""
      full_text = []
      with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                              text = page.extract_text()
                              if text:
                                                full_text.append(text)
                                    return "\n".join(full_text)


def parse_invoice_fields(text, filename):
      """
          Parse common invoice fields from extracted text.
              Customize these patterns to match your client's document format.
                  """
      record = {"file": filename}

    # Example: extract a date (matches formats like 01/15/2024 or 2024-01-15)
      date_match = re.search(r"(\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|\d{4}-\d{2}-\d{2})", text)
      record["date"] = date_match.group(1) if date_match else ""

    # Example: extract a dollar total
      total_match = re.search(r"(?:total|amount due)[^\d]*(\$?[\d,]+\.\d{2})", text, re.IGNORECASE)
      record["total"] = total_match.group(1) if total_match else ""

    # Example: extract invoice number
      inv_match = re.search(r"(?:invoice\s*#?|inv\s*#?)\s*([A-Z0-9-]+)", text, re.IGNORECASE)
      record["invoice_number"] = inv_match.group(1) if inv_match else ""

    # Raw text for client review
      record["raw_text_preview"] = text[:300].replace("\n", " ")

    return record


def process_folder(folder_path, output_file):
      """Walk a folder, extract data from each PDF, and export to CSV."""
      records = []

    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith(".pdf")]

    if not pdf_files:
              print(f"No PDF files found in '{folder_path}'.")
              return

    print(f"Found {len(pdf_files)} PDF(s). Processing...")

    for filename in pdf_files:
              filepath = os.path.join(folder_path, filename)
              try:
                            text = extract_text_from_pdf(filepath)
                            record = parse_invoice_fields(text, filename)
                            records.append(record)
                            print(f"  [OK] {filename}")
except Exception as e:
            print(f"  [ERROR] {filename}: {e}")
            records.append({"file": filename, "error": str(e)})

    df = pd.DataFrame(records)
    df.to_csv(output_file, index=False)
    print(f"\nDone. {len(records)} records saved to '{output_file}'.")


if __name__ == "__main__":
      parser = argparse.ArgumentParser(description="Extract data from a folder of PDFs.")
      parser.add_argument("--folder", default="./invoices", help="Path to PDF folder")
      parser.add_argument("--output", default="output.csv", help="Output CSV filename")
      args = parser.parse_args()

    process_folder(args.folder, args.output)
