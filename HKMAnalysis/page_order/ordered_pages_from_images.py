from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from PIL import Image
import os

def create_pdf(image_folder, output_pdf):
    # Get the list of image files sorted by file name
    image_files = sorted(os.listdir(image_folder))
    print(image_files)

    # Create a PDF canvas
    c = canvas.Canvas(output_pdf, pagesize=letter)

    # Set margin
    margin = 10

    for image_file in image_files:
        if image_file.endswith(('.jpg', '.jpeg', '.png', '.gif')):
            # Load image
            img_path = os.path.join(image_folder, image_file)
            img = Image.open(img_path)

            # Get image dimensions
            width, height = img.size

            # Calculate aspect ratio
            aspect_ratio = width / height

            # Adjust dimensions for fitting within the page
            if aspect_ratio > 1:  # Horizontal image
                new_width = letter[0] - 2 * margin
                new_height = new_width / aspect_ratio
            else:  # Vertical image
                new_height = letter[1] - 2 * margin
                new_width = new_height * aspect_ratio

            # Add the image to the PDF
            c.drawImage(img_path, margin, margin, width=new_width, height=new_height)

            # Add a new page if there are more images
            if image_file != image_files[-1]:
                c.showPage()

    # Save the PDF
    c.save()

# Example usage:
image_folder = "images"
output_pdf = "output.pdf"
create_pdf(image_folder, output_pdf)
