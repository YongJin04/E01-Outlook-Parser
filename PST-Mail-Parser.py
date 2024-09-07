from aspose.email.storage.pst import PersonalStorage, StandardIpmFolder
from aspose.email.mapi import MapiMessage, ContactSaveFormat
from datetime import timezone
import glob
import csv
import os

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
            "delivery_time_datetime": mapi_message.delivery_time if mapi_message.delivery_time else '',
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

if __name__ == "__main__":
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
