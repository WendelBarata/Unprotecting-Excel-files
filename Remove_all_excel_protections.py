'''
                        ATTENTION !!!
At this point the code removes the protection from the VBA Project, however it
corrupts the bits of the file, to access the macros you need to export the
macros and then import them into another excel file.
'''

from zipfile import ZipFile, ZIP_DEFLATED
import os
import tempfile
from lxml import etree
from pathlib import Path


def remove_all_excel_protection(excel_file: Path, ext: str = 'xlsx') -> None:
    """
    Function to remove protection from an Excel file
    (password protected sheets and VBA project).

    This function extracts the contents of the Excel file, removes the
    protection from all sheet files, removes the password from the VBA project,
    and creates a new Excel file without protection.

    Parameters:
        excel_file (Path): The path to the protected Excel file.
        ext (str): The extension of the Excel file, default is 'xlsx'.

    Raises:
        *FileNotFoundError:
        If the 'worksheets' directory is not found in the extracted Excel file.

    Returns:
        None
    """
    with tempfile.TemporaryDirectory() as temp_dir:
        with ZipFile(excel_file, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        # Path to the worksheets
        worksheets_path = os.path.join(temp_dir, 'xl', 'worksheets')

        # Check if the worksheets directory exists
        if not os.path.exists(worksheets_path):
            raise FileNotFoundError("Could not find the 'worksheets' directory"
                                    " in the extracted Excel file.")

        # Get the names of all sheet files
        sheet_files = [f for f in os.listdir(worksheets_path)
                       if f.endswith('.xml')]

        # Remove the protection tag from all sheet files
        for sheet_file in sheet_files:
            file_path = os.path.join(worksheets_path, sheet_file)
            tree = etree.parse(file_path)
            root = tree.getroot()

            # Remove the sheetProtection tag
            sheet_protection = root.find(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheetProtection") # noqa
            if sheet_protection is not None:
                root.remove(sheet_protection)
                print(f"Protection removed from file: {sheet_file}")

            # Write the modified sheet file back to the temporary directory
            tree.write(file_path,
                       xml_declaration=True,
                       encoding='UTF-8',
                       standalone=True)

        # Path to the vbaProject.bin file
        vba_project_path = os.path.join(temp_dir, 'xl', 'vbaProject.bin')

        if os.path.exists(vba_project_path):
            with open(vba_project_path, 'rb') as file:
                vba_content = file.read()

            # Remove the password from the VBA project
            new_vba_content = remove_vba_password_from_bin(vba_content)

            # Write the modified content back to vbaProject.bin
            with open(vba_project_path, 'wb') as file:
                file.write(new_vba_content)
            print("VBA password removed successfully.")

        # Create a new zip file without protection
        unprotected_file = f'{str(excel_file)[:-5]}_unprotected.{ext}'
        with ZipFile(unprotected_file, 'w', ZIP_DEFLATED) as new_zip:
            for foldername, _, filenames in os.walk(temp_dir):
                for filename in filenames:
                    file_path = os.path.join(foldername, filename)
                    new_zip.write(file_path,
                                  os.path.relpath(file_path, temp_dir))

        print(f"Unprotected file saved as: {unprotected_file}")


def remove_vba_password_from_bin(vba_content: bytes) -> bytes:
    """
    Function to remove the password from the VBA project binary content.

    This function searches for specific markers in the binary content and
    replaces the password with null bytes.

    Parameters:
        vba_content (bytes): The binary content of the vbaProject.bin file.

    Returns:
        bytes: The modified binary content with the password removed.
    """
    # Identify and modify all areas in the binary where the password
    # might be stored
    start_strs = [b'DPB', b'DPx', b'DPw']
    new_vba_content = vba_content

    for start_str in start_strs:
        start_index = new_vba_content.find(start_str)
        while start_index != -1:
            print(f"Marker '{start_str.decode()}'"
                  f"found at index: {start_index}")
            end_index = start_index + 6
            new_vba_content = (new_vba_content[:start_index]
                               + b'\x00\x00\x00\x00\x00\x00'
                               + new_vba_content[end_index:])
            start_index = new_vba_content.find(start_str, start_index + 1)

    return new_vba_content


# Path to the Excel file
excel_file = Path(r"C:\Users\wisaias\Downloads\teste_.xlsm")
extension = excel_file.split('.')[-1]


# Removing protection from the Excel file
remove_all_excel_protection(excel_file, extension)
