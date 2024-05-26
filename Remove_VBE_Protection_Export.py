'''
                        ATTENTION !!!
At this point the code removes the protection from the VBA Project, however it
corrupts the bits of the file, to access the macros you need to export the
macros and then import them into another excel file.
'''

import os
from zipfile import ZipFile, ZIP_DEFLATED
import shutil
from pathlib import Path


def remove_vba_project_password(excel_path: Path, ext: str = 'xlsx') -> None:
    """
    Function to remove the VBA project password from an Excel file.

    This function takes a protected Excel file as input, extracts the
    contents of the file, removes the password from the VBA project,
    and creates a new Excel file with the modified content.

    Parameters:
        excel_path (Path): The path to the protected Excel file.
        ext (str): The extension of the Excel file, default is 'xlsx'.

    Raises:
        *FileNotFoundError:
        If the 'vbaProject.bin' file is not found in the extracted Excel file.

    Returns:
        None
    """
    try:
        # Rename the Excel file to a .zip file
        zip_path = str(excel_path).replace(f'.{ext}', '.zip')
        os.rename(excel_path, zip_path)

        # Extract the contents of the zip file
        with ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall('excel_content')

        # Path to the vbaProject.bin file
        vba_project_path = os.path.join(
            'excel_content', 'xl', 'vbaProject.bin')

        # Check if the vbaProject.bin file exists
        if not os.path.exists(vba_project_path):
            raise FileNotFoundError("The 'vbaProject.bin' file was not found.")

        # Read the content of the vbaProject.bin
        with open(vba_project_path, 'rb') as file:
            vba_content = file.read()

        # Remove the password from the VBA Project
        new_vba_content = remove_vba_password_from_bin(vba_content)

        # Write the modified content back to vbaProject.bin
        with open(vba_project_path, 'wb') as file:
            file.write(new_vba_content)

        # Create a new zip file with the modified content
        new_zip_path = zip_path.replace('.zip', '_unprotected.zip')
        with ZipFile(new_zip_path, 'w', ZIP_DEFLATED) as zip_ref:
            for folder_name, subfolders, filenames in os.walk('excel_content'):
                for filename in filenames:
                    file_path = os.path.join(folder_name, filename)
                    arcname = os.path.relpath(file_path, 'excel_content')
                    zip_ref.write(file_path, arcname)
        print(f"New zip file created: {new_zip_path}")

        # Rename the new zip file to .xlsm
        new_excel_path = new_zip_path.replace('.zip', f'.{ext}')
        os.rename(new_zip_path, new_excel_path)

        # Clean up temporary files
        os.rename(zip_path, excel_path)
        shutil.rmtree('excel_content')

    except Exception as e:
        print(f"Error removing VBA password: {e}")
        # Clean up temporary files in case of failure
        if os.path.exists('excel_content'):
            shutil.rmtree('excel_content')
        if os.path.exists(zip_path) and not os.path.exists(excel_path):
            os.rename(zip_path, excel_path)


def remove_vba_password_from_bin(vba_content):
    """
    Function to remove the password from the VBA project binary content.

    This function identifies and modifies all areas in the binary
    where the password might be stored.

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


# Path to the protected Excel file
excel_file = Path(r"C:\Users\wisaias\Downloads\teste_.xlsm")
extension = excel_file.split('.')[-1]

# Call the function
remove_vba_project_password(excel_file, extension)
