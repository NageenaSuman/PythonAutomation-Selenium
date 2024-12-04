import os
import fitz
import tempfile
from easygui import multenterbox
from PIL import Image


# Get user input
fieldNames = ["Enter the value for DPI:", "Enter the PDF file path:", "Output folder for images:"]
fieldValues = list(multenterbox(msg='Fill in values for the fields', title='PDF Resolution I/P in DPI', fields=(fieldNames)))
print(fieldValues)


def extract_and_resize_images(pdf_path, output_folder):
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


            # Save the resized image
            output_path = os.path.join(output_folder, f"page_{page_number + 1}_image_{img_index}.{image_ext}")
            pil_image.save(output_path, dpi=(int(fieldValues[0]), int(fieldValues[0])))

            # Delete the temporary file
            os.unlink(temp_file_path)

    # Close the PDF document
    pdf_document.close()


extract_and_resize_images(fieldValues[1], fieldValues[2])
