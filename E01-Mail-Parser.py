import win32com.client
import pandas as pd
import hashlib
import pytsk3
import pyewf
import glob
import sys
import os

class EWFImgInfo(pytsk3.Img_Info):
    """This class extends pytsk3.Img_Info to support EWF image files."""
    def __init__(self, ewf_handle):
        self._ewf_handle = ewf_handle
        super().__init__(url="")

    def close(self):
        """Closes the EWF image handle."""
        self._ewf_handle.close()

    def read(self, offset, size):
        """Reads data from the EWF image at the specified offset and size."""
        self._ewf_handle.seek(offset)
        return self._ewf_handle.read(size)

    def get_size(self):
        """Returns the total size of the EWF image."""
        return self._ewf_handle.get_media_size()

def get_file_type(filepath):
    """Determines the file type based on file extension, supports E01 and raw images."""
    return 'E01' if filepath.lower().endswith('.e01') else 'raw'

def open_ewf_image(file_paths):
    """Opens an EWF image file for reading using pyewf handle."""
    ewf_handle = pyewf.handle()
    ewf_handle.open(file_paths)
    return ewf_handle

def read_image_file(imgpath, imgtype):
    """Reads an image file, supports both E01 and raw formats."""
    if imgtype == "E01":
        filenames = pyewf.glob(imgpath)
        ewf_handle = open_ewf_image(filenames)
        img_info = EWFImgInfo(ewf_handle)
    else:
        img_info = pytsk3.Img_Info(imgpath)
    return img_info

def format_title_one(title):
    """Formats the title for consistent display length across image file outputs."""
    total_length = 50  # Desired total length of title bar
    left_padding = (total_length - len(title)) // 2
    right_padding = total_length - (left_padding + len(title))
    return ' ' + '=' * left_padding + ' ' + title + ' ' + '=' * right_padding + ' '

def print_all_partitions_with_windows_directory(img_info, output_dir, img_path):
    """Prints partition and user information for partitions containing a Windows directory."""
    title = format_title_one(os.path.basename(img_path))
    print(f'\n{title}')
    try:
        partition_table = pytsk3.Volume_Info(img_info)
        for partition in partition_table:
            try:
                fs = pytsk3.FS_Info(img_info, offset=partition.start * 512)
                if fs.info.ftype == pytsk3.TSK_FS_TYPE_NTFS and has_windows_directory(fs):
                    print(f" Partition Name : {partition.desc.decode()}")
                    extract_count = print_users_directories_with_outlook(fs, output_dir)
                    print(f" Extracted : {extract_count}")
            except Exception as e:
                pass
    except Exception as e:
        print(f" Failed to read partition info: {str(e)}")

def has_windows_directory(fs):
    """Checks if a 'Windows' directory exists in the file system."""
    try:
        root_dir = fs.open_dir(path="/")
        for entry in root_dir:
            if entry.info.meta and entry.info.meta.type == pytsk3.TSK_FS_META_TYPE_DIR:
                if entry.info.name.name.decode().lower() == 'windows':
                    return True
    except Exception as e:
        print(f" Error checking for 'Windows' directory: {str(e)}")
    return False

def print_users_directories_with_outlook(fs, output_dir):
    """Prints Outlook OST file information for each user directory."""
    extracted_files = 0
    try:
        users_dir = fs.open_dir(path="/Users")
        for entry in users_dir:
            if entry.info.meta and entry.info.meta.type == pytsk3.TSK_FS_META_TYPE_DIR:
                dir_name = entry.info.name.name.decode()
                if dir_name not in [".", ".."]:
                    if contains_appdata_directory(fs, f"/Users/{dir_name}"):
                        ost_files = extract_ost_files(fs, f"/Users/{dir_name}/AppData/Local/Microsoft/Outlook", output_dir)
                        pst_files = extract_pst_files(fs, f"/Users/{dir_name}/OneDrive/문서/Outlook Files", output_dir)
                        if ost_files:
                            print(f"    User Name : {dir_name}")
                            print("        (Outlook-OST-Directory O)")
                            for file in ost_files:
                                print(f"            -> {file}")
                            list_outlook_files(fs, dir_name)
                            extracted_files += len(ost_files)
                            extracted_files += len(pst_files)
                        else:
                            print(f"    User Name : {dir_name}")
                            print("        (Outlook-OST-Directory X)")
                            list_outlook_files(fs, dir_name)
                            extracted_files += len(pst_files)
    except Exception as e:
        print(f" Failed to list Users subdirectories: {str(e)}")
    return extracted_files

def contains_appdata_directory(fs, path):
    """Checks if an 'AppData' directory exists within the specified path."""
    try:
        dir_to_check = fs.open_dir(path=path)
        for entry in dir_to_check:
            if entry.info.meta and entry.info.meta.type is pytsk3.TSK_FS_META_TYPE_DIR:
                if entry.info.name.name.decode().lower() == 'appdata':
                    return True
    except Exception as e:
        print(f" Error checking for 'AppData' directory in {path}: {str(e)}")
    return False

def extract_ost_files(fs, path, output_dir):
    """Extracts OST files from the specified Outlook directory."""
    extracted_files = []
    try:
        outlook_dir = fs.open_dir(path=path)
        for entry in outlook_dir:
            if entry.info.meta and entry.info.meta.type is pytsk3.TSK_FS_META_TYPE_REG:
                file_name = entry.info.name.name.decode()
                if file_name.lower().endswith('.ost'):
                    file_path = os.path.join(output_dir, file_name)
                    with open(file_path, 'wb') as f:
                        file_data = entry.read_random(0, entry.info.meta.size)
                        f.write(file_data)
                    extracted_files.append(file_name)
    except Exception as e:
        # Suppress error message printing
        pass
    return extracted_files

def list_outlook_files(fs, dir_name):
    """Lists filenames in the Outlook Files directory for a user."""
    path = f"/Users/{dir_name}/OneDrive/문서/Outlook Files"
    try:
        outlook_files_dir = fs.open_dir(path=path)
        has_files = False
        print("        (Outlook-PST-Directory O)")
        for entry in outlook_files_dir:
            if entry.info.meta and entry.info.meta.type == pytsk3.TSK_FS_META_TYPE_REG:
                file_name = entry.info.name.name.decode()
                print(f"            -> {file_name}")
        has_files = True
                
        if not has_files:
            print("        No files in directory.")
    except Exception as e:
        print("        (Outlook-PST-Directory X)")

def extract_pst_files(fs, path, output_dir):
    """Extracts OST files from the specified Outlook directory."""
    extracted_files = []
    try:
        outlook_dir = fs.open_dir(path=path)
        for entry in outlook_dir:
            if entry.info.meta and entry.info.meta.type is pytsk3.TSK_FS_META_TYPE_REG:
                file_name = entry.info.name.name.decode()
                if file_name.lower().endswith('.pst'):
                    file_path = os.path.join(output_dir, file_name)
                    with open(file_path, 'wb') as f:
                        file_data = entry.read_random(0, entry.info.meta.size)
                        f.write(file_data)
                    extracted_files.append(file_name)
    except Exception as e:
        # Suppress error message printing
        pass
    return extracted_files

def process_image_file(img_path):
    """Processes each image file, reads the file, and extracts OST files."""
    img_type = get_file_type(img_path)
    hasher = hashlib.sha256()
    try:
        with open(img_path, 'rb') as afile:
            buf = afile.read(500 * 1024 * 1024)  # Read in 500MB chunks
            while buf:
                hasher.update(buf)
                buf = afile.read(500 * 1024 * 1024)
    except IOError as e:
        print(f" Unable to open file {img_path}: {str(e)}")
        sys.exit(1)

    hash_value = hasher.hexdigest()
    output_directory_name = os.path.basename(img_path) + '-' + hash_value
    output_directory = os.path.join("./extracted_files", output_directory_name)
    os.makedirs(output_directory, exist_ok=True)

    img_info = read_image_file(img_path, img_type)
    print_all_partitions_with_windows_directory(img_info, output_directory, img_path)

    if img_info and img_type == "E01":
        img_info.close()

def E01_OST_PST_Parser(img_file):
    process_image_file(img_file)
def format_title(filename):
    """Formats the title to include the directory name before '-' and append 'backup.pst'.
    The title is then centered in a line of '=' symbols for clear display in the console."""
    base_name = filename.split('-')[0]  # Extract base name from filename before '-'
    title = f"{base_name}:backup.pst"  # Append ':backup.pst' to the base name for the title
    total_length = 70  # Desired total length of title bar
    left_padding = (total_length - len(title)) // 2  # Calculate left padding for centering
    right_padding = total_length - (left_padding + len(title))  # Calculate right padding for centering
    return '=' * left_padding + title + '=' * right_padding  # Return the formatted title

def ensure_outlook_constants():
    """Ensure that Outlook constants are loaded for script usage, essential for identifying email types."""
    outlook = win32com.client.Dispatch("Outlook.Application")
    win32com.client.gencache.EnsureDispatch('Outlook.Application')  # Force generation of cache for constants
    constants = win32com.client.constants  # Load Outlook constants
    return constants

def connect_to_outlook(pst_file_path):
    """Connect to Outlook using a specified PST file path and add the PST file to the Outlook profile."""
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")  # Get MAPI namespace
    try:
        outlook.AddStore(pst_file_path)  # Attempt to add PST file to the session
        root_folder = outlook.Folders.Item(outlook.Folders.Count)  # Get the last folder (added PST)
        return root_folder
    except Exception as e:
        print(f"  Failed to add PST file: {e}")  # Error handling for PST file addition
        return None

def read_folder_messages(folder, constants):
    """Read and parse email messages from a specified folder using Outlook constants."""
    messages = folder.Items
    data = []
    for message in messages:
        try:
            if message.Class == constants.olMail:  # Filter only email messages
                received_time = getattr(message, 'ReceivedTime', None)  # Safely get the received time
                mail_data = {
                    "mail_index": folder.Name or '',
                    "sender_name": getattr(message, 'SenderName', ''),
                    "sender_account": getattr(message, 'SenderEmailAddress', ''),
                    "receiver_account": getattr(message, 'To', ''),
                    "cc_account": getattr(message, 'CC', ''),
                    "received_time": received_time.strftime('%Y-%m-%d %H:%M:%S') if received_time else '',
                    "title": getattr(message, 'Subject', ''),
                    "body": getattr(message, 'Body', '')[:10000],
                    "attachment_file_name": ', '.join([attachment.FileName for attachment in message.Attachments]) if message.Attachments.Count > 0 else ''
                }
                data.append(mail_data)  # Collect all relevant mail data
        except Exception as e:
            print(f"  Error reading message in '{folder.Name}': {e}")
    return data

def process_all_folders(folders, constants):
    """Recursively process all folders and compile email data."""
    all_data = []
    for folder in folders:
        folder_data = read_folder_messages(folder, constants)
        all_data.extend(folder_data)
        all_data.extend(process_all_folders(folder.Folders, constants))  # Recursively process subfolders
    return all_data

def save_emails_to_csv(emails, filename, base_dir):
    """Save the list of email dictionaries to a CSV file."""
    if emails:
        df = pd.DataFrame(emails)
        df.to_csv(filename, index=False, na_rep='', encoding='utf-8-sig')  # Save DataFrame to CSV
        relative_path = os.path.relpath(filename, start=base_dir)  # Calculate relative path for printing
        print(f"  Saving  : {relative_path} -> Saved  {len(emails)} emails")

def process_pst_files():
    base_dir = os.path.join(os.getcwd(), 'extracted_files')  # Define the base directory for PST files
    pst_files = glob.glob(f"{base_dir}/**/*.pst", recursive=True)  # Find all PST files recursively
    constants = ensure_outlook_constants()  # Load constants used for processing emails

    for pst_file in pst_files:
        pst_basename = os.path.basename(pst_file).split('.')[0]  # Get basename without extension
        directory_name = os.path.basename(os.path.dirname(pst_file))  # Get directory name of the PST file
        print(format_title(directory_name))  # Format and print the directory title
        relative_path = os.path.relpath(pst_file, start=base_dir)  # Get relative path of the PST file
        print(f"  Parsing : {relative_path} ->", end=' ')
        root_folder = connect_to_outlook(pst_file)  # Connect to Outlook and get root folder
        if root_folder:
            emails = process_all_folders([root_folder], constants)  # Process all folders in the PST file
            print(f"Parsed {len(emails)} emails")
            csv_filename = os.path.splitext(pst_basename)[0] + '.csv'  # Create CSV filename
            csv_path = os.path.join(os.path.dirname(pst_file), csv_filename)  # Define full path for CSV
            save_emails_to_csv(emails, csv_path, base_dir)  # Save emails to CSV
        else:
            print("  Could not open the PST file. Please check if it's not corrupted and retry.")

def PST_Mail_Parser():
    process_pst_files()

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: E01-Mail-Parser.exe -<Option Name>")
        sys.exit(1)

    option = sys.argv[1]

    if option not in ["-op", "-mp"]:
        print("Invalid option. Use '-op' or '-mp'.")
        sys.exit(1)

    if option == "-op":
        if len(sys.argv) < 2:
            print("Usage: E01-Mail-Parser.exe -op <image file path> <image file path> ...")
            sys.exit(1)
        try:
            for img_file in sys.argv[2:]:
                E01_OST_PST_Parser(img_file)
        except Exception as e:
            print(f"An error occurred while processing with OutlookParser: {e}")
    elif option == "-mp":
        try:
            PST_Mail_Parser()
        except Exception as e:
            print(f"An error occurred while processing with MailParser: {e}")

