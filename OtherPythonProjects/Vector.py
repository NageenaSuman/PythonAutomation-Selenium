import json
import PyPDF2

# Open the PDF file
pdf_file = open('D:/Temp/Test.pdf', 'rb')

# Create a PDF reader object
pdf_reader = PyPDF2.PdfReader(pdf_file)

# Get the number of pages in the PDF file
num_pages = len(pdf_reader.pages)

# Loop through all the pages and extract the text
for page in range(num_pages):
    page_obj = pdf_reader.pages[page]
    print(page_obj.extract_text())

# Close the PDF file
pdf_file.close()