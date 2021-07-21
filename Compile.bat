@ECHO OFF


RMDIR /Q /S dist


SET "external_imports=C:\Users\Ethan (Personal)\Google Drive\Computer Sync\Software\My_Programs\Pycharm\@Custom_Packages\Word_Doc_Convert\external_imports"

SET external_import_data="%external_imports%;external_imports\."

venv\Scripts\pyinstaller --add-data %external_import_data% --onefile word_to_pdf.py
venv\Scripts\pyinstaller --add-data %external_import_data% --onefile word_to_images.py
venv\Scripts\pyinstaller --add-data %external_import_data% --onefile pdf_to_images.py


RMDIR /Q /S build
RMDIR /Q /S __pycache__
DEL "*.spec"