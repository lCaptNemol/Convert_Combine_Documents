import os
import glob
from pypdf import PdfReader, PdfWriter

# If its your first time using these tools 
    #make sure to instal these dependancies by running this command in your terminal. 
    #pip3 install pypdf
# Place your Word documents (.docx files) inside a subfolder of the base directory:
# Below copy and paste the base directory's path
BASE_DIR = "/Users/km/Documents/Projects/Combine_PDFs"

def list_folders(base_path):
    """Returns a list of subdirectories inside the base path."""
    return [f for f in os.listdir(base_path) if os.path.isdir(os.path.join(base_path, f))]

def get_user_selected_folder(base_path):
    """Prompts the user to select a folder from the available options."""
    folders = list_folders(base_path)
    if not folders:
        print("No subdirectories found in the base directory.")
        return None
    
    print("\nAvailable folders:")
    for i, folder in enumerate(folders, 1):
        print(f"{i}. {folder}")

    while True:
        try:
            choice = int(input("\nEnter the number of the folder to use: "))
            if 1 <= choice <= len(folders):
                return folders[choice - 1]
            else:
                print("Invalid selection. Please choose a valid folder number.")
        except ValueError:
            print("Invalid input. Please enter a number.")

def merge_pdfs(input_folder, output_folder):
    """Merges all PDFs in the input folder into a single PDF file."""
    os.makedirs(output_folder, exist_ok=True)
    
    pdf_files = sorted(glob.glob(os.path.join(input_folder, "*.pdf")))

    if not pdf_files:
        print("No valid PDF files found in the folder.")
        return

    output_filename = "combined_document.pdf"
    output_path = os.path.join(output_folder, output_filename)

    pdf_writer = PdfWriter()

    for pdf in pdf_files:
        print(f"Adding: {pdf}")  # Debugging line
        reader = PdfReader(pdf)
        for page in reader.pages:
            pdf_writer.add_page(page)

    with open(output_path, "wb") as output_pdf:
        pdf_writer.write(output_pdf)

    print(f"Merged PDF saved as: {output_path}")
    print(f"Total number of PDFs combined: {len(pdf_files)}")

if __name__ == "__main__":
    selected_folder = get_user_selected_folder(BASE_DIR)
    
    if selected_folder:
        input_folder = os.path.join(BASE_DIR, selected_folder)
        output_folder = os.path.join(BASE_DIR, "Merged_PDFs")  # Output folder
        merge_pdfs(input_folder, output_folder)
    else:
        print("No folder selected. Exiting program.")