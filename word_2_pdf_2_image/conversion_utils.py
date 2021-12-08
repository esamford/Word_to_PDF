import os
import sys
from io import BytesIO
from typing import List

import pdf2image
import win32com.client
from PIL import Image


def open_file_as_bytes_io(file_path: str) -> BytesIO:
    assert isinstance(file_path, str)
    assert os.path.exists(file_path)
    assert os.path.isfile(file_path)

    with open(file_path, 'rb') as file:
        file_bytes = BytesIO(file.read())
        file_bytes.seek(0)
    return file_bytes


def convert_word_to_pdf(word_path: str or BytesIO, pdf_path: str, **kwargs):
    """
    :param word_path: The file path to the Word document.
    :param pdf_path: The file path to the new PDF document.

    NOTE: This will only work on Windows OS. Multiplatform functionality has not been set up yet.
    """
    assert isinstance(word_path, str)
    assert isinstance(pdf_path, str)
    assert os.path.exists(word_path)
    assert os.path.isfile(word_path)

    # Convert to absolute paths to prevent raising an exception later.
    word_path = os.path.abspath(word_path)
    pdf_path = os.path.abspath(pdf_path)

    pdf_format_code = 17

    try:
        word_object = win32com.client.Dispatch('Word.Application')
        doc_object = word_object.Documents.Open(word_path)
        doc_object.SaveAs(pdf_path, FileFormat=pdf_format_code)
        doc_object.Close()
        word_object.Quit()
    except Exception as ex:
        msg = "An exception has occurred while converting a Word document to a PDF document. If you are running " + \
              "on a platform that is not Windows, this is probably the reason why this exception was raised. " \
              f"\nException details: \n\n{ex}"
        raise Exception(msg)

    assert os.path.exists(pdf_path)
    if "delete_word_doc" in kwargs and kwargs["delete_word_doc"]:
        os.remove(word_path)


def convert_pdf_to_images(pdf_path: str) -> List[Image.Image]:
    """
    :param pdf_path: The file path to the new PDF document.
    :return: A list of PIL Image.Image objects; each image is for a page in the PDF.
    """
    assert isinstance(pdf_path, str)
    assert os.path.exists(pdf_path)
    assert os.path.isfile(pdf_path)

    # The program acts differently between PyInstaller --onefile and --onedir.
    try:
        working_directory = sys._MEIPASS
    except AttributeError:
        working_directory = os.getcwd()

    poppler_path = os.path.join(working_directory, "resources", "poppler", "Library", "bin")
    if not os.path.exists(poppler_path):
        raise Exception(
            "Cannot find path to Poppler resource, which is used to convert PDF pages to images. \n"
            "Please check that it was packaged correctly during the PyInstaller compile process."
        )
    return pdf2image.convert_from_path(
        pdf_path, poppler_path=poppler_path,
        dpi=300  # DPI of 300 is about the max someone can notice on the best printers.
    )
