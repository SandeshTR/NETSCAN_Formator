from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Cm , Inches
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml.etree import QName
import re
import comtypes.client
import os
import time
# from devCode import FormatTableStyle, RemoveTable , RemoveSectionBreaks, CheckSmallCapsFunction
# from devCode import replace_bullets 
import logging

   


# Precompile the regular expressions
email_pattern = re.compile(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b')
url_pattern = re.compile(r'\b(?:https?://|www\.)\S+\b')
doc_link_pattern = re.compile(r'\b\S+\.(com|gov|aspx|html|txt|rtf|pdf|doc|docx|xls|xlsx)\b')


def convert_file_to_docx(file_path, output_path):
    """ 
    Convert a PDF or DOC file to DOCX format using Microsoft Word. 
    """
    word = comtypes.client.CreateObject('Word.Application')
    word.Visible = False      # Keep Word application hidden

    try:
        logging.info(f"Converting {file_path} to {output_path}")
        doc = word.Documents.Open(file_path)
        doc.SaveAs(output_path, FileFormat=16)     # FileFormat=16 is for .docx
        doc.Close()
        logging.info(f"Conversion successful: {output_path}")

    except Exception as e:
        logging.info(f"An error occurred during conversion: {e}")
    finally:
        word.Quit()



def format_document(doc_path,region_code = None):
    """
    Load a DOCX document, apply formatting and modifications, and save it.
    """
    # Load the document
    doc = Document(doc_path)

    # Apply formatting to all paragraphs
    for paragraph in doc.paragraphs:
        format_paragraph(paragraph)
        align_paragraph(paragraph)
        apply_bullet_formatting(paragraph)

    # Apply formatting to all tables
    for table in doc.tables:
        indent_table_left(table)
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    format_paragraph(paragraph)
                    align_paragraph(paragraph)
                    apply_bullet_formatting(paragraph)

    # remove_spacing_between_paras(doc)
    # remove_tabs_from_document(doc)
    
    # Save the modified document
    doc.save(doc_path)
    logging.info(f" Formatted document and saved: {doc_path}")

    # custom call functions
    # FormatTableStyle.add_border_doc(doc_path,doc_path)        #adds border to all the table elements
    # RemoveTable.replace_image_with_text(doc_path,"{Non Displayable Image}",doc_path)       # replaces image with "{Non Displayable Image}"
    # RemoveTable.remove_line_numbering(doc_path, doc_path)   
    # replace_bullets.process_bullets_in_document(doc_path)   
    # replace_bullets.set_bullet_follow_char_to_space(doc_path) 
    # replace_bullets.set_bullet_font(doc_path)    
    # # remove_header_footer_sections(doc_path,region_code)
    # CheckSmallCapsFunction.check_small_caps(doc_path)
    # RemoveSectionBreaks.remove_section_breaks(doc_path)
    word_file_indentation(doc_path)                             
    
def is_excluded_text(text):
    return email_pattern.search(text) or url_pattern.search(text) or doc_link_pattern.search(text)


def format_paragraph(paragraph):
    """
    Apply specific formatting to a paragraph. Remove hyperlink text (but not images), apply text font and style.
    """
    replacements = {
        '–': '~#8211;',
        '§': '~#167;',
        ' ,': ',',
        '°': '~#176;',
        '•': '~#8226;',
        '✓': '/',
        '☒': '[x]',
        '→': '-->'
    }

    # # Accumulate text from all runs to process leading spaces
    # full_text = ''.join(run.text for run in paragraph.runs)
    # stripped_text = full_text.lstrip()

    # # Determine the amount of leading whitespace removed
    # leading_spaces_removed = len(full_text) - len(stripped_text)

    # # Update runs with the stripped text
    # remaining_text = stripped_text
    # for run in paragraph.runs:
    #     # Check if the run contains an image
    #     if run._element.findall('.//w:drawing', namespaces=paragraph._element.nsmap):
    #         continue  # Skip this run if it contains an image

    #     if leading_spaces_removed > 0:
    #         run_text_length = len(run.text)
    #         if run_text_length <= leading_spaces_removed:
    #             run.text = ''
    #             leading_spaces_removed -= run_text_length
    #         else:
    #             run.text = run.text[leading_spaces_removed:]
    #             leading_spaces_removed = 0
    #     if remaining_text:
    #         run.text = remaining_text[:len(run.text)]
    #         remaining_text = remaining_text[len(run.text):]

    # # Replace text according to the replacements dictionary
    # for old_text, new_text in replacements.items():
    #     if old_text in paragraph.text:
    #         inline = paragraph.runs
    #         for run in inline:
    #             run.text = run.text.replace(old_text, new_text)

    for run in paragraph.runs:
        # Check if the run contains an image
        if run._element.findall('.//w:drawing', namespaces=paragraph._element.nsmap):
            continue  # Skip this run if it contains an image

        if 'HYPERLINK' in run.element.xml:
            run.clear()  # Remove the hyperlink text

        if is_excluded_text(run.text):
            run.font.underline = False  # Remove underline if the text matches the exclusion criteria

        # replacing starting " with " and ending " with " and replacing all ' with ' 
        text = run.text
        text = re.sub(r'"(\b)', r'“\1', text)
        text = re.sub(r'(\b)"', r'”\1', text)
        text = re.sub(r'"([.,!?;:)\s]|$)', r'”\1', text)
        text = text.replace("'", "’")
        run.text = text

        # run.font.underline = False  # This will make all the underlines disappear
        run.font.size = Pt(12)
        run.font.name = 'Times New Roman'
        run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')



def align_paragraph(paragraph):
    """
    Align the paragraph and set indentation.
    """
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.space_before = Pt(3)
    paragraph.paragraph_format.space_after = Pt(3)
    paragraph.paragraph_format.left_indent = Cm(0)
    paragraph.paragraph_format.first_line_indent = Cm(0)

def apply_bullet_formatting(paragraph):
    """
    Apply bullet formatting if the paragraph is part of a bullet list.
    """
    if paragraph._element.xpath('.//w:numPr'):
        paragraph.paragraph_format.left_indent = Cm(0)
        paragraph.paragraph_format.first_line_indent = Cm(0)

def indent_table_left(table):
    """
    Indent the entire table to the left with 0 cm.
    """
    tbl = table._element
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.append(tblPr)
    tblInd = tblPr.xpath('w:tblInd')
    if tblInd:
        tblPr.remove(tblInd[0])
    tblInd = OxmlElement('w:tblInd')
    tblInd.set(qn('w:w'), '0')
    tblInd.set(qn('w:type'), 'dxa')
    tblPr.append(tblInd)


def add_nac_to_section(doc, paragraph):
    """
    Add "NAC" to sections within the document.
    """
    section_pattern = re.compile(r'(Section|Sec\.?)\s*\d+\.\s*NAC\s*\d+[A-Z]?\.?\d+\s+is hereby amended to read as follows:')

    for i in range(len(doc.paragraphs) - 1):  # Iterate to the second-to-last paragraph
        paragraph = doc.paragraphs[i]
        if section_pattern.search(paragraph.text):
            next_paragraph = doc.paragraphs[i + 1]
            if not next_paragraph.text.strip().startswith("NAC"):
                next_paragraph.runs[0].text = "NAC " + next_paragraph.runs[0].text


# def remove_header_footer_sections(doc_path,region_code):
#     # Open the document
#     doc = Document(doc_path)

#     # Remove headers and footers
#     for section in doc.sections:
#         # Remove header references
#         section.different_first_page_header_footer = False
#         section.header.is_linked_to_previous = True
        
#         # Remove footer references
#         section.footer.is_linked_to_previous = True

#         # Clear any existing header and footer content
#         section.header.paragraphs.clear()
#         if region_code is not None:
#             if region_code.strip().lower() != 'ca':        
#                 section.footer.paragraphs.clear()

#     # Remove header and footer references
#     for section in doc.sections:
#         sectPr = section._sectPr
#         for child in sectPr:
#             if child.tag.endswith(('headerReference', 'footerReference')):
#                 sectPr.remove(child)

#     # Set header and footer distance to 0
#     for section in doc.sections:
#         section.header_distance = 0
#         section.footer_distance = 0

#     # Save the modified document
#     doc.save(doc_path)
#     logging.info(f"Header and footer sections removed from the document: {doc_path}")


def clean_up_temp_file(temp_file_path):
    """
    Remove the temporary DOCX file if it exists.
    """
    if os.path.exists(temp_file_path):
        os.remove(temp_file_path)



# def remove_spacing_between_paras(doc):
#     """
#     This function is used to remove spacing between paragraphs while preserving content
#     """
#     logging.info('removed spacing between paras')
#     for para in doc.paragraphs:
#         # Check if the paragraph is empty and doesn't contain any of these elements
#         if (not para.text.strip()  and 
#             not para._element.findall('.//w:drawing', namespaces=para._element.nsmap) and
#             not para._element.findall('.//w:ins', namespaces=para._element.nsmap) and
#             not para._element.findall('.//w:del', namespaces=para._element.nsmap) and
#             not para._element.findall('.//w:smartTag', namespaces=para._element.nsmap)):
            
#             if not para._element.getparent().tag.endswith('tbl'):
#                 p = para._element
#                 p.getparent().remove(p)
#                 p._element = p._p = None


# def remove_tabs_from_document(doc):
#     """Function to remove tabs from the file text while preserving images"""

#     def process_paragraph(paragraph):
#         for run in paragraph.runs:
#             if not run.element.findall('.//w:drawing', namespaces=run.element.nsmap):
#                 # Only process runs that don't contain images
#                 run.text = run.text.replace('\t', ' ')

#     # Process paragraphs
#     for paragraph in doc.paragraphs:
#         process_paragraph(paragraph)

#     # Process tables
#     for table in doc.tables:
#         for row in table.rows:
#             for cell in row.cells:
#                 for paragraph in cell.paragraphs:
#                     process_paragraph(paragraph)

#     logging.info(f"Tabs removed successfully")


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



def check_file_type_and_convert(input_file_path: str, output_doc_path: str):
    """
    Check the file type and convert it to DOCX format, then apply formatting.
    """
    if os.path.exists(output_doc_path):
        os.remove(output_doc_path)

    base_name, file_extension = os.path.splitext(input_file_path)
    temp_doc_path = f"{base_name}.docx"

    if file_extension.strip().lower() in ['.doc', '.pdf', '.docx']:
        convert_file_to_docx(input_file_path, temp_doc_path)
    else:
        logging.warning("Unsupported file type. Only .doc and .pdf are supported.")
        return None

    time.sleep(5)  # Ensure the file is fully converted before proceeding
    return temp_doc_path

# format_document(r"C:\File\NETSCAN\Input\NETSCAN_CO_Test_14\Exception\2025-00021_co_p.docx")