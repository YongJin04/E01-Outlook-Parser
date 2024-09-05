import win32com.client
import pandas as pd
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
                    "mail_index": folder.Name,
                    "sender_name": message.SenderName,
                    "sender_account": message.SenderEmailAddress,
                    "receiver_account": message.To,
                    "cc_account": message.CC,
                    "received_time": received_time.strftime('%Y-%m-%d %H:%M:%S') if received_time else 'N/A',
                    "title": message.Subject,
                    "body": message.Body[:2000],
                    "attachment_file_name": ', '.join([attachment.FileName for attachment in message.Attachments]) if message.Attachments.Count > 0 else 'None'
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

def save_emails_to_csv(emails, filename='emails.csv'):
    """Save the list of email dictionaries to a CSV file."""
    if emails:
        df = pd.DataFrame(emails)
        df.to_csv(filename, index=False, encoding='utf-8-sig')
        print(f"Saved {len(emails)} emails to {filename}")
        # Call the function to delete tmp files after saving the CSV

def main(pst_file_path):
    if not os.path.exists(pst_file_path):
        print(f"File does not exist: {pst_file_path}")
        return

    constants = ensure_outlook_constants()
    root_folder = connect_to_outlook(pst_file_path)
    if root_folder:
        print(f"Successfully connected to {pst_file_path}")
        emails = process_all_folders([root_folder], constants)
        print(f"Parsed {len(emails)} emails.")
        save_emails_to_csv(emails)
    else:
        print("Could not open the PST file. Please check if it's not corrupted and retry.")

if __name__ == "__main__":
    pst_file_path = input("Enter the path to your PST file: ")
    
    if not os.path.exists(pst_file_path):
        print(f"File does not exist: {pst_file_path}")
        exit()

    constants = ensure_outlook_constants()
    root_folder = connect_to_outlook(pst_file_path)
    if root_folder:
        print(f"Successfully connected to {pst_file_path}")
        emails = process_all_folders([root_folder], constants)
        print(f"Parsed {len(emails)} emails.")
        save_emails_to_csv(emails)
    else:
        print("Could not open the PST file. Please check if it's not corrupted and retry.")

