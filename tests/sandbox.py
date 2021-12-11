import os

import utils
from word_2_pdf_2_image import conversion_utils

word_path = os.path.join("Test_Documents", "! Test_Document.docx")
pdf_path = utils.get_pdf_path(word_path)
if os.path.exists(pdf_path):
    os.remove(pdf_path)

conversion_utils.convert_word_to_pdf(word_path, pdf_path)
