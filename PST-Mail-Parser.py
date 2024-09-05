import win32com.client
import pandas as pd
import glob
import os

def ensure_outlook_constants():
    """Ensure that Outlook constants are loaded for script usage."""
    outlook = win32com.client.Dispatch("Outlook.Application")
    win32com.client.gencache.EnsureDispatch('Outlook.Application')
    constants = win32com.client.constants
    return constants

def connect_to_outlook(pst_file_path):
    """Connect to Outlook using a specified PST file path."""
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    try:
        outlook.AddStore(pst_file_path)
        root_folder = outlook.Folders.Item(outlook.Folders.Count)
        return root_folder
    except Exception as e:
        print(f"Failed to add PST file: {e}")
        return None

def read_folder_messages(folder, constants):
    """Read and parse email messages from a specified folder."""
    messages = folder.Items
    data = []
    for message in messages:
        try:
            if message.Class == constants.olMail:
                received_time = getattr(message, 'ReceivedTime', None)
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
                data.append(mail_data)
        except Exception as e:
            print(f"Error reading message in '{folder.Name}': {e}")
    return data

def process_all_folders(folders, constants):
    """Recursively process all folders and compile email data."""
    all_data = []
    for folder in folders:
        folder_data = read_folder_messages(folder, constants)
        all_data.extend(folder_data)
        all_data.extend(process_all_folders(folder.Folders, constants))
    return all_data

def save_emails_to_csv(emails, filename):
    """Save the list of email dictionaries to a CSV file."""
    if emails:
        df = pd.DataFrame(emails)
        df.to_csv(filename, index=False, na_rep='', encoding='utf-8-sig')  # Using na_rep to replace NaN with empty string
        print(f"Saved {len(emails)} emails to {filename}")

def process_pst_files():
    base_dir = os.path.join(os.getcwd(), 'extracted_files')
    pst_files = glob.glob(f"{base_dir}/**/*.pst", recursive=True)
    constants = ensure_outlook_constants()

    for pst_file in pst_files:
        print(f"Processing: {pst_file}")
        root_folder = connect_to_outlook(pst_file)
        if root_folder:
            emails = process_all_folders([root_folder], constants)
            csv_filename = os.path.splitext(os.path.basename(pst_file))[0] + '.csv'
            save_emails_to_csv(emails, os.path.join(os.path.dirname(pst_file), csv_filename))
            print(f"Parsed {len(emails)} emails.")
        else:
            print("Could not open the PST file. Please check if it's not corrupted and retry.")

if __name__ == "__main__":
    process_pst_files()
