import comtypes.client
import logging


def convert_file_to_docx(file_path, output_path):
    """ 
    Convert a PDF or DOC file to DOCX format using Microsoft Word. 
    """
    word = comtypes.client.CreateObject('Word.Application')
    word.Visible = False                                          # Keep Word application hidden

    try:
        logging.info(f"Converting {file_path} to {output_path}")
        doc = word.Documents.Open(file_path)
        doc.SaveAs(output_path, FileFormat=16)                    # FileFormat=16 is for .docx
        doc.Close()
        logging.info(f"Conversion successful: {output_path}")

    except Exception as e:
        logging.info(f"An error occurred during conversion: {e}")
    finally:
        word.Quit()