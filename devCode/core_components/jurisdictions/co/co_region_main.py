from docx import Document
from docx.oxml.ns import qn
import pytesseract
from PIL import Image
import io, os
from docx.oxml import OxmlElement
from core_components.jurisdictions.co import co_redline
from core_components.jurisdictions.co import co_aft
from logs.logs_handler import get_logger
from common_func.word_paragraph_operations import convert_hyperlink_to_text

logger = get_logger(__name__)

# logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(filename)s - %(message)s')


def extract_images_with_locations(input_docx_path,header_check=False):
    logger.info(f"header check : {header_check}")
 
    doc = Document(input_docx_path)
    images_with_locations = []
    index = 0
    logger.info('read document')
    doc_part_to_consider = doc.paragraphs[:11] if header_check else doc.paragraphs  #depending on weather you want to check header or all the values use this
     

    for i, paragraph in enumerate(doc_part_to_consider):
        for element in paragraph._element.iter():
            if element.tag == qn('w:drawing'):
                blip = element.find('.//a:blip', namespaces={'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'})
                if blip is not None:
                    img_id = blip.get(qn('r:embed'))
                    image_part = doc.part.related_parts[img_id]
                    images_with_locations.append((index, i, image_part.blob))
                    index += 1

    logger.info(f'number of images identified {len(images_with_locations)}')
    return images_with_locations



def extract_text_from_images(images_with_locations):
    """extract images and there location from docx """
    texts = []
    for idx, location, img_blob in images_with_locations:
        image = Image.open(io.BytesIO(img_blob))
        text = pytesseract.image_to_string(image=image)
        print(f'text extracted is : {text}')
        if text.strip():                       #filtering and adding only non empty image locations
            texts.append((idx, location, text))
            logger.info(f'added location text {location}')
    return texts


def insert_single_text_in_paragraph(input_docx_path, location, text):
    """This function adds logo text present at the top of page along with 
    file type for ADDITIONAL INFORMATION/ EMERGENCY JUSTIFICATION/ BASIS AND PURPOSE files"""
    
    base_name, file_extension = os.path.splitext(input_docx_path)
    doc = Document(input_docx_path)
                                 #only one logo present in docx file
    paragraph = doc.paragraphs[location]
    paragraph.clear()
    logger.info(" cleared text - removed logo content ")
    for run in paragraph.runs:
        if '_' in run.text:
            run.text = run.text.replace('_', '')
    if location < 3:
        if 'addinfo' in base_name.strip().lower():
            paragraph.add_run('ADDITIONAL INFORMATION \n' + text.strip().replace('\n\n','\n')).bold = True
        elif 'emergency' in base_name.strip().lower():
            paragraph.add_run('EMERGENCY JUSTIFICATION \n' + text.strip().replace('\n\n','\n')).bold = True    
        elif 'basisandpurpose' in base_name.strip().lower():
            paragraph.add_run('BASIS AND PURPOSE\n'+ text.strip().replace('\n\n','\n')).bold = True         
        else:
            paragraph.add_run(text.strip().replace('\n\n','\n')).bold = True

    doc.save(input_docx_path)
    logger.info(f"Modified document saved as {input_docx_path}")




def insert_multiple_text_in_paragraph(input_docx_path,img_location_list):
    """Function to add multiple text based on the list index position provided"""
    doc = Document(input_docx_path)
    
    sorted_list = sorted(img_location_list,key=lambda x: x[1],reverse=True)
    logger.info(f'sorted list {sorted_list}')

    for index,location,content in sorted_list:
        if "colorado" in content.lower():
            paragraph = doc.paragraphs[location]

            # Remove all runs (including images) from the paragraph
            for run in paragraph.runs:
                run.clear()
            
            paragraph.add_run(content.strip().replace('\n\n', '\n')).bold = True
            logger.info(f'Added colorado data {content}')
        else:
            logger.info(f"Skipping content at location {location} as it doesn't contain 'Colorado'")
    
    base_name = os.path.basename(input_docx_path).strip().lower()

    if any(name in base_name.strip().lower() for name in ['addinfo','emergency','basisandpurpose']):
        paragraph = doc.paragraphs[0].insert_paragraph_before()
        if 'addinfo' in base_name.strip().lower():
            paragraph.add_run('ADDITIONAL INFORMATION').bold = True
        elif 'emergency' in base_name.strip().lower():
            paragraph.add_run('EMERGENCY JUSTIFICATION').bold = True
        elif 'basisandpurpose' in base_name.strip().lower():
            paragraph.add_run('BASIS AND PURPOSE').bold = True

    
    # Save the modified document
    doc.save(input_docx_path)      


def add_first_line_header(input_docx_path):
    base_name, file_extension = os.path.splitext(input_docx_path)

    doc = Document(input_docx_path)   
    
    paragraph = doc.paragraphs[0].insert_paragraph_before()

    if 'addinfo' in base_name.strip().lower():
        if doc.paragraphs[0].text =='ADDITIONAL INFORMATION' or doc.paragraphs[1].text == 'ADDITIONAL INFORMATION' or doc.paragraphs[2].text == 'ADDITIONAL INFORMATION':
            return
        paragraph.add_run('ADDITIONAL INFORMATION').bold = True
    elif 'emergency' in base_name.strip().lower():
        paragraph.add_run('EMERGENCY JUSTIFICATION').bold = True    
    elif 'basisandpurpose' in base_name.strip().lower():
        paragraph.add_run('BASIS AND PURPOSE').bold = True         
    doc.save(input_docx_path)
    logger.info(f"Added first line info when images count is zero ,document saved as {input_docx_path}")



def main_co_files(input_docx_path,source_file_path):
    """Main function for co region, this function assigns and calls functions based """
    
    logger.info(f"Starting main_co_files function with input: {input_docx_path}")
    text_inserted = False
    #get basename of docx
    base_name, file_extension = os.path.splitext(input_docx_path)

    #logic to check initial header/logo present and add text at that location
    images_with_locations = extract_images_with_locations(input_docx_path,header_check=False)
    
    logger.info(f"Number of images extracted: {len(images_with_locations)}")

    if len(images_with_locations) > 0:  # File contains headers
        img_texts = extract_text_from_images(images_with_locations)
        texts = [text for text in img_texts  if text[2]]
        logger.debug(f"Extracted texts: {texts}")


        if len(texts)>0:
            if len(images_with_locations) > 1 and len(texts) > 1:
                # This is for files which has logo to begin with
                logger.info("calling multiple text insert")
                if any('colorado' in text[2].strip().lower() for text in texts):
                    insert_multiple_text_in_paragraph(input_docx_path=input_docx_path,img_location_list=img_texts )
            else:
                logger.info("calling single text insert")
                #need to check what basename is aft or kinda 
                insert_single_text_in_paragraph(input_docx_path=input_docx_path,location=texts[0][1],text=texts[0][2])                
                text_inserted = True

    #check file type and perform action
    if 'aft' in base_name.strip().lower():
        #chek pdf present as source and if keywords present
        co_aft.determine_aft_file_type(input_docx_path=input_docx_path,source_file_path=source_file_path)

    elif 'redline' in base_name.strip().lower():
        #calling redline code
        co_redline.main(input_docx_path) 

    elif 'addinfo' in base_name.strip().lower():   #Added extra check for addinfo as no data was getting added (added on 2025-01-16)
        if not text_inserted:
            # add_first_line_header(input_docx_path)
            if len(images_with_locations) > 0 and len(texts) > 0:
                insert_single_text_in_paragraph(input_docx_path=input_docx_path,location=0,text=texts[0][2])

    if len(images_with_locations) <= 0:
        add_first_line_header(input_docx_path=input_docx_path)

    # # run common guideline logics
    # replace_bullets_with_text(input_docx_path)
    convert_hyperlink_to_text(input_docx_path)     # Issue Cause is this function which is not working properly for linked footers
    # remove_tabs_from_document(input_docx_path)
    