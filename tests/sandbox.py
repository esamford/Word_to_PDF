import os

import conversion_utils
import utils

word_path = os.path.join("Test_Documents", "Sentence.docx")
pdf_path = utils.get_pdf_path(word_path)

conversion_utils.convert_word_to_pdf(word_path, pdf_path)
