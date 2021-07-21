import sys

from conversion_utils import convert_pdf_to_images
import utils

if __name__ == "__main__":
    for pos, arg in enumerate(sys.argv):
        # Skip the first position that represents this executable.
        if pos == 0:
            continue

        try:
            utils.print_status(pos, len(sys.argv) - 1, arg)

            extension = arg[arg.rfind('.'):]
            if extension == ".pdf":
                images = convert_pdf_to_images(arg)
                image_names = utils.get_image_paths(arg, len(images))
                assert len(image_names) == len(images)
                for image, name in zip(images, image_names):
                    image.save(name)
        except Exception as ex:
            with open("Exception.log", "w") as file:
                file.write(str(ex))
