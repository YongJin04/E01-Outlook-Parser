from aspose.email.storage.pst import PersonalStorage
from datetime import timezone
import hashlib
import pytsk3
import pyewf
import glob
import sys
import csv
import os

# ==================== E01_to_ost_and_pst ==================== #

class EWFImgInfo(pytsk3.Img_Info):
    def __init__(self, ewf_handle):
        self._ewf_handle = ewf_handle
        super().__init__(url="")

    def close(self):
        self._ewf_handle.close()

    def read(self, offset, size):
        self._ewf_handle.seek(offset)
        return self._ewf_handle.read(size)

    def get_size(self):
        return self._ewf_handle.get_media_size()

def get_file_type(filepath):
    return 'E01' if filepath.lower().endswith('.e01') else 'raw'

def open_ewf_image(file_paths):
    ewf_handle = pyewf.handle()
    ewf_handle.open(file_paths)
    return ewf_handle

def read_image_file(imgpath, imgtype):
    if imgtype == "E01":
        filenames = pyewf.glob(imgpath)
        ewf_handle = open_ewf_image(filenames)
        img_info = EWFImgInfo(ewf_handle)
    else:
        img_info = pytsk3.Img_Info(imgpath)
    return img_info

def format_title(title):
    total_length = 50
    left_padding = (total_length - len(title)) // 2
    right_padding = total_length - (left_padding + len(title))
    return ' ' + '=' * left_padding + ' ' + title + ' ' + '=' * right_padding + ' '

def print_all_partitions_with_windows_directory(img_info, output_dir, img_path):
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
    extracted_files = 0
    try:
        users_dir = fs.open_dir(path="/Users")
        for entry in users_dir:
            if entry.info.meta and entry.info.meta.type == pytsk3.TSK_FS_META_TYPE_DIR:
                dir_name = entry.info.name.name.decode()
                if dir_name not in [".", ".."]:
                    if contains_appdata_directory(fs, f"/Users/{dir_name}"):
                        ost_files = extract_files(fs, f"/Users/{dir_name}/AppData/Local/Microsoft/Outlook", output_dir, '.ost')
                        pst_files = extract_files(fs, f"/Users/{dir_name}/OneDrive/문서/Outlook Files", output_dir, '.pst')
                        if ost_files:
                            print(f"    User Name : {dir_name}")
                            print("        (Outlook-OST-Directory O)")
                            for file in ost_files:
                                print(f"            - {file}")
                            list_outlook_files(fs, dir_name, output_dir)

                            extracted_files += len(ost_files)
                            extracted_files += len(pst_files)
    except Exception as e:
        print(f" Failed to list Users subdirectories: {str(e)}")
    return extracted_files

def contains_appdata_directory(fs, path):
    try:
        dir_to_check = fs.open_dir(path=path)
        for entry in dir_to_check:
            if entry.info.meta and entry.info.meta.type is pytsk3.TSK_FS_META_TYPE_DIR:
                if entry.info.name.name.decode().lower() == 'appdata':
                    return True
    except Exception as e:
        print(f" Error checking for 'AppData' directory in {path}: {str(e)}")
    return False

def extract_files(fs, path, output_dir, extension):
    extracted_files = []
    try:
        outlook_dir = fs.open_dir(path=path)
        for entry in outlook_dir:
            if entry.info.meta and entry.info.meta.type is pytsk3.TSK_FS_META_TYPE_REG:
                file_name = entry.info.name.name.decode()
                if file_name.lower().endswith(extension):
                    file_path = os.path.join(output_dir, file_name)
                    with open(file_path, 'wb') as f:
                        file_data = entry.read_random(0, entry.info.meta.size)
                        f.write(file_data)
                    extracted_files.append(file_name)
    except Exception as e:
        pass
    return extracted_files

def list_outlook_files(fs, dir_name, output_dir):
    path = f"/Users/{dir_name}/OneDrive/문서/Outlook Files"
    try:
        outlook_files_dir = fs.open_dir(path=path)
        has_files = False
        print("        (Outlook-PST-Directory O)")
        for entry in outlook_files_dir:
            if entry.info.meta and entry.info.meta.type == pytsk3.TSK_FS_META_TYPE_REG:
                file_name = entry.info.name.name.decode()
                print(f"            - {file_name}", end="")

                full_path = os.path.join(output_dir, file_name)
                pst_to_csv('.\\' + full_path[2:])
        has_files = True
                
        if not has_files:
            print("        No files in directory.")
    except Exception as e:
        print("        (Outlook-PST-Directory X)")

def E01_to_ost_and_pst(img_path):
    img_type = get_file_type(img_path)
    hasher = hashlib.sha256()
    try:
        with open(img_path, 'rb') as afile:
            buf = afile.read(500 * 1024 * 1024)
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
        
# ======================== pst_to_csv ======================== #

def load_pst_messages(pst, folder_name):
    folder = pst.root_folder.get_sub_folder(folder_name)
    if folder is None:
        return []
    return folder.get_contents()

def create_csv_for_pst(pst, pst_file, messages_info):
    csv_filename = f"{os.path.splitext(pst_file)[0]}.csv"
    with open(csv_filename, 'w', newline='', encoding='utf-8-sig') as csvfile:
        fieldnames = ["folder_name", "sender_email", "sender_name", "receiver_emails", "cc_emails", "delivery_time", "subject", "body", "attachments"]
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        for folder_name, messages in messages_info.items():
            display_message_info(messages, pst, folder_name, writer)
    csv_base_name = os.path.basename(csv_filename)
    print(f" -> {csv_base_name}")
    return csv_filename

def display_message_info(messages, pst, folder_name, writer):
    for message_info in messages:
        mapi_message = pst.extract_message(message_info)
        receiver_emails = mapi_message.display_to.split(';') if mapi_message.display_to else ['']
        email_data = {
            "folder_name": folder_name,
            "sender_email": mapi_message.sender_email_address if mapi_message.sender_email_address else '',
            "sender_name": (format_kor_name(mapi_message.sender_name if mapi_message.sender_name else '')).replace(" ", ""),
            "receiver_emails": "; ".join(receiver_emails).strip(),
            "cc_emails": mapi_message.display_cc if mapi_message.display_cc else '',
            "delivery_time": int(mapi_message.delivery_time.strftime("%Y%m%d%H%M%S")) if mapi_message.delivery_time else '',
            "subject": mapi_message.subject if mapi_message.subject else '',
            "body": mapi_message.body[:2000] if mapi_message.body else '',
            "attachments": ", ".join([attachment.display_name for attachment in mapi_message.attachments]) if mapi_message.attachments else ''
        }
        writer.writerow(email_data)

def format_kor_name(name):
    parts = name.split()
    if len(parts) != 2:
        return name
    if len(parts[0]) == 1 and len(parts[1]) == 2:
        return name
    elif len(parts[0]) == 2 and len(parts[1]) == 1:
        return parts[1] + ' ' + parts[0]
    else:
        return name

def pst_to_csv(pst_file):
    with PersonalStorage.from_file(pst_file) as pst:
        folder_names = ["Inbox", "Outbox", "Sent Items", "Deleted Items", "Drafts", "Junk Email"]
        messages_info = {}
        for folder_name in folder_names:
            messages = load_pst_messages(pst, folder_name)
            messages_info[folder_name] = messages
        create_csv_for_pst(pst, pst_file, messages_info)

# ==================== merge_and_sort_csv_files ==================== #

def merge_and_sort_csv_files(directory):
    csv_files = glob.glob(os.path.join(directory, '**', '*.csv'), recursive=True)
    all_data = []
    fieldnames = ["folder_name", "sender_email", "sender_name", "receiver_emails", "cc_emails", "delivery_time", "subject", "body", "attachments"]
    num_files_merged = len(csv_files)
    
    for csv_file in csv_files:
        with open(csv_file, 'r', newline='', encoding='utf-8-sig') as file:
            reader = csv.DictReader(file)
            for row in reader:
                all_data.append(row)

    all_data.sort(key=lambda x: int(x['delivery_time']) if x['delivery_time'] else 0)

    merged_filename = os.path.join(".", 'extract.csv')
    with open(merged_filename, 'w', newline='', encoding='utf-8-sig') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        for data in all_data:
            writer.writerow(data)

    print(f"\n{num_files_merged} CSV files merged and sorted into '{merged_filename}'\n")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: E01-Mail-Parser.exe <E01 file path 1> <E01 file path 2> ...")
        sys.exit(1)
    
    for img_file in sys.argv[1:]:
        E01_to_ost_and_pst(img_file)
    
    merge_and_sort_csv_files(os.path.join(".", "extracted_files"))
    
