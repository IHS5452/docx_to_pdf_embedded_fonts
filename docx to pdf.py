from docx2pdf import convert
import os

def convert_to_pdf_with_embedded_fonts(save_location, output_filename, docx_file):
    output_filepath = os.path.join(save_location, output_filename)
    
    try:
        convert(docx_file, output_filepath, keep_active=False, silent=True, flags="--embed-all-fonts")
        print("Conversion to PDF completed successfully. Output file:", output_filepath)
    except Exception as e:
        print("Error converting to PDF:", str(e))

# Usage example
save_location = input("Enter the save location: ").strip('\'"')
output_filename = input("Enter the output file name: ").strip('\'"')
docx_file = input("Enter the path to the DOCX file: ").strip('\'"')

convert_to_pdf_with_embedded_fonts(save_location, output_filename, docx_file)
