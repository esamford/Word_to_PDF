import os
import sys

import utils
from word_2_pdf_2_image.conversion_utils import convert_word_to_pdf, convert_pdf_to_images

if __name__ == "__main__":
    if len(sys.argv) == 1:
        utils.print_drag_and_drop_message()
    else:
        for pos, arg in enumerate(sys.argv):
            # Skip the first position that represents this executable.
            if pos == 0:
                continue

            try:
                utils.print_status(pos, len(sys.argv) - 1, arg)

                pdf_path = utils.get_pdf_path(arg)
                pdf_exists = os.path.exists(pdf_path)

                extension = arg[arg.rfind('.'):]
                if extension == ".docx" or extension == ".doc":
                    # Convert the Word document to PDF.
                    convert_word_to_pdf(arg, pdf_path)
                    arg = pdf_path

                extension = arg[arg.rfind('.'):]
                if extension == ".pdf":
                    # Get the page images from the PDF.
                    images = convert_pdf_to_images(pdf_path)
                    image_names = utils.get_image_paths(pdf_path, len(images))

                    # Delete the PDF if it did not already exist.
                    while not pdf_exists and os.path.exists(pdf_path):
                        try:
                            os.remove(pdf_path)
                        except Exception:
                            pass

                    # Save the images.
                    assert len(image_names) == len(images)
                    for image, name in zip(images, image_names):
                        image.save(name)
            except Exception as ex:
                with open("Exception.log", "w") as file:
                    file.write(str(ex))
