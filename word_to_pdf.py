import sys

import utils
from word_2_pdf_2_image.conversion_utils import convert_word_to_pdf

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

                extension = arg[arg.rfind('.'):]
                if extension == ".docx" or extension == ".doc":
                    pdf_path = utils.get_pdf_path(arg)
                    convert_word_to_pdf(arg, pdf_path)
            except Exception as ex:
                with open("Exception.log", "w") as file:
                    file.write(str(ex))
