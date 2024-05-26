from zipfile import ZipFile
import shutil
import os
from lxml import etree
from pathlib import Path


def remove_excel_protection(excel_file: Path,
                            ext: str = 'xlsx'
                            ) -> None:
    """
    Function to remove the protection from an Excel file

    This function takes an Excel file as input and removes the protection
    from the file. It does so by extracting the contents of the Excel file,
    removing the sheetProtection tag from all the sheet files, and then
    creating a new Excel file with the modified contents.

    Parameters:
    - excel_file (Path): The path to the Excel file to be unprotected
    - ext (str): The extension of the new Excel file (default is 'xlsx')

    Returns:
    - None: The function does not return any value, it creates a new Excel
      file with the modified contents
    """
    # Create a temporary directory to extract the contents of the Excel file
    with ZipFile(excel_file, 'r') as zip_ref:
        zip_ref.extractall('temp_excel')

    # Get the names of all the sheet files
    sheet_files = [f for f in os.listdir('temp_excel/xl/worksheets')
                   if f.endswith('.xml')]

    # Remove the sheetProtection tag from all the sheet files
    for sheet_file in sheet_files:
        file_path = os.path.join('temp_excel/xl/worksheets', sheet_file)
        tree = etree.parse(file_path)
        root = tree.getroot()

        # Remove the sheetProtection tag from the sheet file
        for elem in root.iter():
            if 'sheetProtection' in elem.tag:
                root.remove(elem)

        # Write the modified sheet file back to the temporary directory
        tree.write(file_path,
                   xml_declaration=True,
                   encoding='UTF-8',
                   standalone=True)

    with ZipFile(f'{excel_file[:-5]}_unprotected.{ext}', 'w') as new_zip:
        for foldername, _, filenames in os.walk('temp_excel'):
            for filename in filenames:
                file_path = os.path.join(foldername, filename)
                new_zip.write(file_path, os.path.relpath(file_path,
                                                         'temp_excel'))

    shutil.rmtree('temp_excel')


arquivo_excel = r"C:\Users\wisaias\Downloads\teste_.xlsm"
extension = "xlsm"
remove_excel_protection(arquivo_excel, extension)
