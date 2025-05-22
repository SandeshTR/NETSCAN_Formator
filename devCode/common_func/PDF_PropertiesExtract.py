import PyPDF2
import pikepdf

def check_pdf_properties(filename):

    with open(filename, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        info = reader.metadata
        
    # Get properties
        producer = info.producer if hasattr(info, 'producer') else "Not available"
        creator = info.creator if hasattr(info, 'creator') else "Not available"
        pdf_version = reader.pdf_header
    
        #print(f"Producer: {producer}")
        #print(f"Creator: {creator}")
        #print(f"PDF Version: {pdf_version}")
        producer_str = str(producer).lower() if producer else ""
        creator_str = str(creator).lower() if creator else ""
    
        if 'word' in producer_str or 'word' in creator_str:
            return "Yes"
        else:
            return "No"

#filename=r"C:\Users\6116591\OneDrive - Thomson Reuters Incorporated\Documents\NETSCAN_Abbyy\Original Input\co_p001aft183.pdf"
#check_pdf_properties(filename)

#with pikepdf.open(filename) as pdf:
#    #print(dir(pdf))
#    metadata = pdf.docinfo
#    producer = metadata.get('/Producer', "Not available")
#    creator = metadata.get('/Creator', "Not available")
#    pdf_version = f"{pdf.pdf_version}.{pdf.pdf_version}"
#    
#    print(f"Producer: {producer}")
#    print(f"Creator: {creator}")
#    print(f"PDF Version: {pdf_version}")
#
#    word_found_pypdf = check_word_in_metadata(producer, creator)
#    print(f"Contains 'word': {word_found_pypdf}")