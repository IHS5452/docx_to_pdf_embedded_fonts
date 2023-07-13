import os
import win32com.client

def convert_to_pdf_with_embedded_fonts(save_location, output_filename, docx_file):
    output_filepath = os.path.join(save_location, output_filename)
    
    try:
        word = win32com.client.Dispatch("Word.Application")
        doc = word.Documents.Open(docx_file)
        doc.SaveAs(output_filepath, FileFormat=17)
        doc.Close()
        word.Quit()
        
        print("Conversion to PDF completed successfully. Output file:", output_filepath)
    except Exception as e:
        print("Error converting to PDF:", str(e))


save_location = input("Enter the save location: ").strip('\'"')
output_filename = input("Enter the output file name: ").strip('\'"')
docx_file = input("Enter the path to the DOCX file: ").strip('\'"')

convert_to_pdf_with_embedded_fonts(save_location, output_filename, docx_file)
