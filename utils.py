import os
from pathlib import Path

import docx
import docx.document


def get_pdf_path(path: str):
    assert isinstance(path, str)
    pdf_path = path[:path.rfind('.')] + ".pdf"
    return pdf_path


def get_image_paths(path: str, num_images: int):
    assert isinstance(path, str)
    assert isinstance(num_images, int)
    assert num_images > 0

    result = []
    base_path = path[:path.rfind('.')]
    if num_images == 1:
        result.append("{}.png".format(base_path))
    else:
        for x in range(num_images):
            result.append("{} ({}).png".format(base_path, x + 1))
    return result


def print_status(current_num: int, total: int, file_path: str):
    assert isinstance(current_num, int)
    assert isinstance(total, int)
    assert isinstance(file_path, str)

    os.system('cls' if os.name == 'nt' else 'clear')
    msg = \
"""
Processing file {} of {}...

{}
""".format(current_num, total, os.path.basename(file_path))
    print(msg)
