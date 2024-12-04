import os
import fitz
import tempfile
import easygui
from PIL import Image

# window title
title = "PDF Resolution I/P in DPI"

# Get user input
myVar = easygui.enterbox("Enter the value for DPI:", title)
print(myVar)


def set_image_dpi(image):
    """
    Rescaling image to 300dpi without resizing
    :param image: An image
    :return: A rescaled image
    """
    image_resize = image
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
    temp_filename = temp_file.name
    image_resize.save(temp_filename, dpi=(int(myVar), int(myVar) ))
    return temp_filename


def extract_and_resize_images(pdf_path, output_folder):
    print(myVar)
    # Create the output folder if it doesn't exist
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Open the PDF file
    pdf_document = fitz.open(pdf_path)

    # Iterate through each page in the PDF
    for page_number in range(len(pdf_document)):
        page = pdf_document.load_page(page_number)

        # Get the page dimensions
        # page_width = int(page.rect.width)
        # page_height = int(page.rect.height)

        # Iterate through the images on the page
        for img_index, img in enumerate(page.get_images(full=True)):
            # Extract the image data
            xref = img[0]
            base_image = pdf_document.extract_image(xref)
            image_bytes = base_image["image"]
            image_ext = base_image["ext"]

            # Save the image to a temporary file
            with tempfile.NamedTemporaryFile(delete=False, suffix=f".{image_ext}") as temp_file:
                temp_file.write(image_bytes)
                temp_file_path = temp_file.name

            # Open the temporary image file using PIL
            pil_image = Image.open(temp_file_path)

            # Resize the image with the specified DPI

            # Save the resized image
            output_path = os.path.join(output_folder, f"page_{page_number + 1}_image_{img_index}.{image_ext}")
            pil_image.save(output_path, dpi=(int(myVar), int(myVar)))

            # Delete the temporary file
            os.unlink(temp_file_path)

    # Close the PDF document
    pdf_document.close()


extract_and_resize_images("D:\Temp\ST.pdf", "D:\Temp\output")
