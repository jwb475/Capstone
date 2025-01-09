import os
from PyPDF2 import PdfReader, PdfWriter

def process_pdf(input_path, output_path):
    reader = PdfReader(input_path)
    writer = PdfWriter()
    
    # Copy all pages except the first two and the last one
    for page in reader.pages:
        # Remove headers and footers by cropping the page
        page.mediabox.upper_right = (
            page.mediabox.right,
            page.mediabox.top - 10  # Adjust this value as needed
        )
        page.mediabox.lower_left = (
            page.mediabox.left,
            page.mediabox.bottom + 10  # Adjust this value as needed
        )
        writer.add_page(page)
    
    # Write the modified PDF
    with open(output_path, "wb") as output_file:
        writer.write(output_file)

def process_folder(folder_path, output):
    for filename in os.listdir(folder_path):
        if filename.lower().endswith('.pdf'):
            input_path = os.path.join(folder_path, filename)
            
            # Remove 'processed_' from the filename if it exists
            if filename.startswith('processed_'):
                new_filename = filename[len('processed_'):]
            else:
                new_filename = filename

            output_path = os.path.join(output, new_filename)
            process_pdf(input_path, output_path)
            print(f"Processed: {filename} -> Saved as: {new_filename}")


# Usage
folder_path = r"C:\Users\Jack\Desktop\Capstone\call"
output = r"C:\Users\Jack\Desktop\Capstone\call"
process_folder(folder_path, output)
