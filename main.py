import PyPDF2
from docx import Document
from docx.enum.text import WD_BREAK

# Open the PDF file in binary mode
with open('document.pdf', 'rb') as pdf_file:
    # Create a PDF reader object
    pdf_reader = PyPDF2.PdfReader(pdf_file)

    # Create a new Word document
    docx_document = Document()

    # Extract text from each page of the PDF and add as a new page in Word document
    for page_num in range(len(pdf_reader.pages)):
        page = pdf_reader.pages[page_num]
        text = page.extract_text()
        docx_document.add_paragraph(text)
        if page_num < len(pdf_reader.pages) - 1:
            docx_document.add_page_break()

    # Save the Word document
    docx_document.save('document.docx')

print("Text from PDF has been extracted and saved to document.docx with each page as a new page in Word document.")
