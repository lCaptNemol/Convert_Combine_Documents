#Pip3 install docx2pdf

import os
import glob
from docx2pdf import convert

# Set base directory
BASE_DIR = "/Users/km/Documents/Projects/Convert_WordDocs"
OUTPUT_FOLDER = os.path.join(BASE_DIR, "Converted_WordDocs")

def convert_all_docs_in_folder(input_folder, output_folder):
    """Converts all .docx files in the input folder to PDFs in the output folder."""
    
    # Ensure the output folder exists
    os.makedirs(output_folder, exist_ok=True)

    # Find all Word documents in the folder (excluding temporary files)
    word_files = [f for f in glob.glob(os.path.join(input_folder, "*.docx")) if not os.path.basename(f).startswith("~$")]

    if not word_files:
        print("No valid Word documents found in the folder.")
        return

    for docx_file in word_files:
        # Define output PDF filename
        pdf_filename = os.path.splitext(os.path.basename(docx_file))[0] + ".pdf"
        output_pdf = os.path.join(output_folder, pdf_filename)

        print(f"Converting: {docx_file} → {output_pdf}")
        
        # Convert the document
        convert(docx_file, output_pdf)

    print(f"\nAll documents converted successfully! PDFs saved in: {output_folder}")

if __name__ == "__main__":
    convert_all_docs_in_folder(BASE_DIR, OUTPUT_FOLDER)