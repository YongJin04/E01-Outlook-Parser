from aspose.email.storage.pst import PersonalStorage, StandardIpmFolder
from aspose.email.mapi import MapiMessage, ContactSaveFormat
from datetime import timezone
import win32com.client
from tqdm import tqdm
import pandas as pd
import hashlib
import pytsk3
import pyewf
import glob
import time
import sys
import csv
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

def format_title(title):
    """Formats the title for consistent display length across image file outputs."""
    total_length = 50  # Desired total length of title bar
    left_padding = (total_length - len(title)) // 2
    right_padding = total_length - (left_padding + len(title))
    return ' ' + '=' * left_padding + ' ' + title + ' ' + '=' * right_padding + ' '

def print_all_partitions_with_windows_directory(img_info, output_dir, img_path):
    """Prints partition and user information for partitions containing a Windows directory."""
    title = format_title(os.path.basename(img_path))
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
        
# =============================================================== #

def load_pst_messages(pst, folder_name):
    folder = pst.root_folder.get_sub_folder(folder_name)
    if folder is None:
        return []
    return folder.get_contents()

def display_message_info(messages, pst, folder_name, writer):
    for message_info in messages:
        mapi_message = pst.extract_message(message_info)  # Extract MapiMessage
        # Extract and handle multiple recipients if any
        receiver_emails = mapi_message.display_to.split(';') if mapi_message.display_to else ['']
        
        # Collect email data for CSV output
        email_data = {
            "folder_name": folder_name,
            "sender_email": mapi_message.sender_email_address if mapi_message.sender_email_address else '',
            "receiver_emails": "; ".join(receiver_emails).strip(),  # Join back for consistent CSV output, or handle individually
            "cc_emails": mapi_message.display_cc if mapi_message.display_cc else '',
            "bcc_emails": mapi_message.display_bcc if mapi_message.display_bcc else '',
            # "delivery_time_datetime": mapi_message.delivery_time if mapi_message.delivery_time else '',
            "delivery_time_unixtime": int(mapi_message.delivery_time.replace(tzinfo=timezone.utc).timestamp()) if mapi_message.delivery_time else '',
            "subject": mapi_message.subject if mapi_message.subject else '',
            "body": mapi_message.body[:200] if mapi_message.body else '',
            "attachments": ", ".join([attachment.display_name for attachment in mapi_message.attachments]) if mapi_message.attachments else ''
        }
        writer.writerow(email_data)

def find_pst_files(directory):
    # Search for all PST files in the specified directory
    return glob.glob(os.path.join(directory, '**', '*.pst'), recursive=True)

def create_csv_for_pst(pst_file, messages_info):
    csv_filename = f"{os.path.splitext(pst_file)[0]}.csv"
    with open(csv_filename, 'w', newline='', encoding='utf-8-sig') as csvfile:  # utf-8-sig for proper encoding in Excel
        fieldnames = ["folder_name", "sender_email", "receiver_emails", "cc_emails", "delivery_time_datetime", "delivery_time_unixtime", "subject", "body", "attachments"]
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        for folder_name, messages in messages_info.items():
            display_message_info(messages, pst, folder_name, writer)
    pst_base_name = os.path.basename(pst_file)
    csv_base_name = os.path.basename(csv_filename)
    print(f"CSV File Created : {os.path.dirname(csv_filename)}\\{pst_base_name} -> {csv_base_name}")  # Updated print statement
    return csv_filename

def parse_mail():
    directory_path = os.path.join(".", "extracted_files")
    pst_files = find_pst_files(directory_path)
    total_emails = 0  # Initialize counter for total emails
    csv_files_created = 0  # Counter for CSV files

    for pst_file in pst_files:
        with PersonalStorage.from_file(pst_file) as pst:
            folder_names = ["Inbox", "Outbox", "Sent Items", "Deleted Items", "Drafts", "Junk Email"]
            messages_info = {}
            for folder_name in folder_names:
                messages = load_pst_messages(pst, folder_name)
                messages_info[folder_name] = messages
                total_emails += len(messages)  # Sum up all messages
            csv_file = create_csv_for_pst(pst_file, messages_info)  # Create a CSV file for each PST
            csv_files_created += 1

    print(f"\nTotal Number Of CSV Files Created / Parsed Emails: {csv_files_created}, {total_emails}\n")



if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: E01-Mail-Parser.exe <E01 file path 1> <E01 file path 2> ...")
        sys.exit(1)
    
    for img_file in sys.argv[1:]:
        process_image_file(img_file)
    
    parse_mail()
