import win32com.client
import pandas as pd
import glob
import os

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

if __name__ == "__main__":
    process_pst_files()  # Entry point to process all PST files
