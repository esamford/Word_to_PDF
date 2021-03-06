@ECHO OFF

RMDIR /Q /S dist

REM Get files/folders to explicitly import using the --add-data argument.
SET data_word_2_pdf_2_image="word_2_pdf_2_image;word_2_pdf_2_image\."
SET data_lxml="venv\Lib\site-packages\lxml;lxml\."
SET data_pythoncom="venv\Lib\site-packages\pywin32_system32;."


venv\Scripts\pyinstaller ^
    --add-data %data_word_2_pdf_2_image% ^
    --add-data %data_pythoncom% ^
    --add-data %data_lxml% ^
    --name "Word_to_PDF" ^
    --onefile word_to_pdf.py
venv\Scripts\pyinstaller ^
    --add-data %data_word_2_pdf_2_image% ^
    --add-data %data_pythoncom% ^
    --add-data %data_lxml% ^
    --name "Word_or_PDF_to_Images" ^
    --onefile word_or_pdf_to_images.py

RMDIR /Q /S build
RMDIR /Q /S __pycache__
DEL "*.spec"
