import win32com.client as win32

def remove_page_numbers_from_headers_and_footers(doc_path):
 
    word = win32.Dispatch("Word.Application")
    word.Visible = False  

 
    doc = word.Documents.Open(doc_path)

    # Loop through all sections in the document  and Clear the header text
    for section in doc.Sections:
        
        for header in section.Headers:
            if header.Exists:
                header.Range.Text = '' 
            
        # Process footers and Clear the footer text
        for footer in section.Footers:
            if footer.Exists:
                footer.Range.Text = ''  

 
    doc.Save()
    doc.Close()

    # Quit Word application
    word.Quit()

 



def remove_repeated_page_numbers(doc_path):
 
    word = win32.Dispatch("Word.Application")
    word.Visible = False   

 
    doc = word.Documents.Open(doc_path)
    
     
    # Adjust the pattern as needed
    page_number_patterns = ["Page ", "PageNumber"]  

    # Loop through each paragraph and check for page number patterns
    for para in doc.Paragraphs:
        text = para.Range.Text
        if any(pattern in text for pattern in page_number_patterns):
            # Replace page number text with an empty string
            for pattern in page_number_patterns:
                if pattern in text:
                    para.Range.Text = text.replace(pattern, "").strip()

     
    doc.Save()
    doc.Close()

   
    word.Quit()

 
#doc_path = r"C:\File\Image\Colorado\Data\co_p004aft059_Part1 - Copy - Copy (2).docx"
#remove_page_numbers_from_headers_and_footers(doc_path)
#remove_repeated_page_numbers(doc_path)

