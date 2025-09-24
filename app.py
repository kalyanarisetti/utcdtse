# Install necessary packages quietly in the Colab environment
import subprocess
import sys

def install(package):
    """A helper function to install a package using pip."""
    subprocess.check_call([sys.executable, "-m", "pip", "install", package, "-q"])

# Install all required libraries
install("python-docx")
install("openpyxl")
install("python-pptx")
install("markdownify")
install("beautifulsoup4")

# Import the main libraries required for the conversion tasks
import os
from google.colab import files
import docx
import openpyxl
from pptx import Presentation
from markdownify import markdownify as md
import zipfile
import io

def convert_file_to_text(file_content, file_name):
    """
    Identifies the file type from its name and converts its content to plain text.
    It handles DOCX, XLSX, PPTX, HTML, and ZIP files.
    """
    # Get the file extension (e.g., '.docx') to determine the file type
    _, file_extension = os.path.splitext(file_name)
    output_text = ""

    try:
        # --- Process Microsoft Word (.docx) files ---
        if file_extension == '.docx':
            # Open the document from in-memory bytes
            doc = docx.Document(io.BytesIO(file_content))
            # Extract text from each paragraph
            output_text = '\n'.join([para.text for para in doc.paragraphs])

        # --- Process Microsoft Excel (.xlsx) files ---
        elif file_extension == '.xlsx':
            # Load the workbook from in-memory bytes
            workbook = openpyxl.load_workbook(io.BytesIO(file_content))
            # Iterate through each sheet and each row to extract cell values
            for sheet in workbook.worksheets:
                for row in sheet.iter_rows():
                    # Join cell values with a tab, handling empty cells
                    output_text += '\t'.join([str(cell.value or '') for cell in row]) + '\n'

        # --- Process Microsoft PowerPoint (.pptx) files ---
        elif file_extension == '.pptx':
            # Load the presentation from in-memory bytes
            prs = Presentation(io.BytesIO(file_content))
            # Extract text from shapes on each slide
            for slide in prs.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text"):
                        output_text += shape.text + '\n'

        # --- Process HTML (.html, .htm) files ---
        elif file_extension in ['.html', '.htm']:
            # Use the markdownify library to convert HTML into clean Markdown text
            # Decode the byte content into a string for processing
            output_text = md(file_content.decode('utf-8', errors='ignore'))

        # --- Process ZIP archives (.zip) ---
        elif file_extension == '.zip':
            # Open the zip file from in-memory bytes
            with zipfile.ZipFile(io.BytesIO(file_content)) as z:
                # Loop through each file within the archive
                for info in z.infolist():
                    # Skip directories and system files (like those from macOS)
                    if not info.is_dir() and not info.filename.startswith('__MACOSX'):
                        with z.open(info.filename) as member_file:
                            member_content = member_file.read()
                            output_text += f"--- Converted content from: {info.filename} ---\n"
                            # Recursively call this function to process the file inside the zip
                            output_text += convert_file_to_text(member_content, info.filename) + "\n\n"
        
        # --- Handle unsupported file types ---
        else:
            output_text = f"Unsupported file type: '{file_extension}'"

    except Exception as e:
        # Return an error message if any part of the conversion fails
        return f"An error occurred while processing {file_name}: {str(e)}"

    return output_text

def main():
    """
    The main function that orchestrates the file upload, conversion, 
    preview display, and download functionality in Google Colab.
    """
    print("Please upload a file to convert (.docx, .xlsx, .pptx, .html, .zip)")
    
    # Trigger Colab's built-in file upload interface
    uploaded = files.upload()

    # Check if a file was actually uploaded
    if not uploaded:
        print("\nNo file was selected. Please run the cell again to upload.")
        return

    # The upload result is a dictionary; get the name and content of the first file
    file_name = next(iter(uploaded))
    file_content = uploaded[file_name]

    print(f"\nProcessing '{file_name}'...")
    
    # Call the conversion function to get the full text
    full_text = convert_file_to_text(file_content, file_name)

    # --- Display Preview of the Output ---
    print("\n--- Conversion Preview (First 1000 characters) ---")
    print(full_text[:1000])
    print("--------------------------------------------------\n")

    # --- Offer the Full Text for Download ---
    # Create a new filename for the output text file
    output_filename = os.path.splitext(file_name)[0] + ".txt"
    # Write the converted text to a local file in the Colab environment
    with open(output_filename, "w", encoding="utf-8") as f:
        f.write(full_text)

    print(f"Conversion complete. Preparing '{output_filename}' for download.")
    # Trigger Colab's file download utility
    files.download(output_filename)

# --- Execute the main function ---
if __name__ == "__main__":
    main()

