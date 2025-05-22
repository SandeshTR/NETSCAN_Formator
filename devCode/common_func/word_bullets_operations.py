from docx.shared import Pt, Cm , Inches
import os, gc
import win32com.client as win32
import re
from docx.oxml.ns import qn
from docx import Document
import logging
# from docx.oxml.shared import qn
from docx.oxml import OxmlElement
import comtypes.client as com

# logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(filename)s - %(message)s')

def process_bullets_in_document(file_path, replacement_text="~#8226;") -> None:
    """
    Process bullets in the document and replace them with a custom character.
    """
    # Initialize the Word application
    word_app = com.CreateObject('Word.Application')
    word_app.Visible = False
 
    # Open the document
    doc = word_app.Documents.Open(file_path)
    paragraphs = doc.Paragraphs
    total_paragraphs = paragraphs.Count
 
    try:
        for paragraph in doc.Paragraphs:
            try:
                list_format = paragraph.Range.ListFormat
                if list_format.ListType == 0 and paragraph.Range.Text.strip().startswith('•'):
                    find_text='•'
                    range = paragraph.Range
                    find = range.Find
                    find.Text = find_text
                    find.Replacement.Text = replacement_text
                    find.Execute(Replace=2)
                    
                if list_format.ListType in [2, 4]:
                    if list_format.ListType == 2 or (
                        list_format.ListType == 4 and
                        not re.match(
                            # r'^(\(?[a-zA-Z0-9]+[\)\.]\)?(\([a-zA-Z0-9]+\))?|[a-zA-Z0-9]+\.[a-zA-Z0-9]+|[Pp]art\s+[a-zA-Z0-9]+[\.:\-\s]*)$',
                            r'^(?!.*([Ss]ection\s+\d+|[Hh]eading\s+\d+|[Cc]hapter\s+\d+|[Aa]rticle\s+\d+))(\d+(\.\d+)*|\(?[a-zA-Z0-9]+[\)\.]\)?(\([a-zA-Z0-9]+\))?|[a-zA-Z]+\.[a-zA-Z0-9]+|[Pp]art\s+[a-zA-Z0-9]+[\.:\-\s]*|[a-zA-Z0-9]+\.[a-zA-Z0-9]+|[a-zA-Z]+\s+\d+|\d+\..*)$',
                            list_format.ListString
                        )
                    ):
                        list_format.RemoveNumbers()
                        paragraph.Range.InsertBefore(replacement_text + " ")
 
                    if paragraph.SpaceAfter != 0:
                        paragraph.SpaceAfter = 0
                    if paragraph.SpaceBefore != 0:
                        paragraph.SpaceBefore = 0
                    if paragraph.LineSpacing != 12:
                        paragraph.LineSpacing = 12
            except Exception as e:
                logging.error(f"Error processing paragraph: {e}")
            finally:
                
                del paragraph
    finally:
        
        doc.Save()
        doc.Close()
        word_app.Quit()
        gc.collect()
        logging.info(f"Replaced {total_paragraphs} bullets in {os.path.basename(file_path)}")

def set_bullet_font(doc_path):
    '''Setting the bullet font to Times New Roman and size to 12pt'''
    doc = Document(doc_path)

    try:
        numbering_part0 = doc.part.numbering_part
    except:
        logging.warning(f"No numbering part found in {os.path.basename(doc_path)}. Skipping bullet font setting.")
        return
    
    del numbering_part0
    # Find all numbering definitions in the document
    numbering = doc.part.numbering_part.numbering_definitions._numbering

    # Iterate through all abstract numbering definitions
    for abstract_num in numbering.findall(qn('w:abstractNum')):
        for lvl in abstract_num.findall(qn('w:lvl')):
            # Find the run properties for the level
            rPr = lvl.find(qn('w:rPr'))
            if rPr is None:
                rPr = OxmlElement('w:rPr')
                lvl.append(rPr)

            # Set font to Times New Roman
            rFonts = rPr.find(qn('w:rFonts'))
            if rFonts is None:
                rFonts = OxmlElement('w:rFonts')
                rPr.append(rFonts)
            rFonts.set(qn('w:ascii'), 'Times New Roman')
            rFonts.set(qn('w:hAnsi'), 'Times New Roman')
            rFonts.set(qn('w:cs'), 'Times New Roman')

            # Set font size to 12pt
            sz = rPr.find(qn('w:sz'))
            if sz is None:
                sz = OxmlElement('w:sz')
                rPr.append(sz)
            sz.set(qn('w:val'), '24')  # 24 half-points = 12 points
    
    doc.save(doc_path)
    logging.info(f"Set Bullet font to Times New Roman size 12 in {os.path.basename(doc_path)}")

def set_bullet_follow_char_to_space(doc_path):
    """
    Sets the follow character of list items [Bullets] to a space in a Word document.
    """
    doc = Document(doc_path)

    try:
        numbering_part = doc.part.numbering_part
    except:
        logging.warning(f"No numbering part found in {os.path.basename(doc_path)}. Skipping bullet follow character setting.")
        return

    numbering_part = doc.part.numbering_part
    if numbering_part is None:
        return

    try:
        numbering = numbering_part.numbering_definitions._element
    except AttributeError:
        numbering = numbering_part.element

    for abstract_num in numbering.findall(qn('w:abstractNum')):
        for lvl in abstract_num.findall(qn('w:lvl')):
            suff = lvl.find(qn('w:suff'))
            if suff is None:
                suff = OxmlElement('w:suff')
                suff.set(qn('w:val'), 'space')
                lvl.append(suff)
            elif suff.get(qn('w:val')) != 'space':
                suff.set(qn('w:val'), 'space')

            ind = lvl.find(qn('w:ind'))
            if ind is None:
                ind = lvl.find(qn('w:ind'))
                if ind is None:
                    ind = OxmlElement('w:ind')
                    ind.set(qn('w:left'), '0')
                    ind.set(qn('w:hanging'), '0')
                    lvl.append(ind)
                else:
                    ind.set(qn('w:left'), '0')
                    ind.set(qn('w:hanging'), '0')
            else:
                ind.set(qn('w:left'), '0')
                ind.set(qn('w:hanging'), '0')
    doc.save(doc_path)
    logging.info(f"Set Bullet follow character to space in {os.path.basename(doc_path)}")

def apply_bullet_formatting(paragraph):
    """
    Apply bullet formatting if the paragraph is part of a bullet list.
    """
    if paragraph._element.xpath('.//w:numPr'):
        paragraph.paragraph_format.left_indent = Cm(0)
        paragraph.paragraph_format.first_line_indent = Cm(0)