import hashlib
import pytsk3
import pyewf
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

def format_title(title):
    """Formats the title for consistent display length across image file outputs."""
    total_length = 50  # Desired total length of title bar
    left_padding = (total_length - len(title)) // 2
    right_padding = total_length - (left_padding + len(title))
    return ' ' + '=' * left_padding + ' ' + title + ' ' + '=' * right_padding + ' '

def print_all_partitions_with_windows_directory(img_info, output_dir, img_path):
    """Prints partition and user information for partitions containing a Windows directory."""
    title = format_title(os.path.basename(img_path))
    print(title)
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
                        if ost_files:
                            print(f"    User Name : {dir_name} (Outlook O)")
                            for file in ost_files:
                                print(f"        -> {file}")
                            extracted_files += len(ost_files)
                        else:
                            print(f"    User Name : {dir_name} (Outlook X)")
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

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python script.py <E01 file path 1> <E01 file path 2> ...")
        sys.exit(1)
    
    for img_file in sys.argv[1:]:
        process_image_file(img_file)
