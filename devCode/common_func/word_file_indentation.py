from docx import Document
import logging
from docx.shared import Inches


def word_file_indentation(doc_path):
    """ This function is used to set margins and selectively add first-line indentation to the word file """
    doc = Document(doc_path)
    
    # Loop through all sections and set margins to 1 inch
    for section in doc.sections:
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)
    
    # Loop through all paragraphs
    for paragraph in doc.paragraphs:
        # Remove any existing indentation
        paragraph.paragraph_format.left_indent = Inches(0)
        paragraph.paragraph_format.right_indent = Inches(0)
        paragraph.paragraph_format.first_line_indent = Inches(0)
    
    doc.save(doc_path)
    logging.info(f"Margins adjusted and selective first-line indentation applied to the document: {doc_path}")