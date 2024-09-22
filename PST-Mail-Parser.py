from aspose.email.storage.pst import PersonalStorage
from datetime import timezone, timedelta
import glob
import sys
import csv
import os
import re

# ======================== pst_to_csv ======================== #

def get_source_account(pst):
    messages = load_pst_messages(pst, "Sent Items")
    
    if not messages:
        messages = load_pst_messages(pst, "보낸 편지함")
    
    if not messages:
        return None
    
    first_message_info = messages[0]
    first_message = pst.extract_message(first_message_info)

    return first_message.sender_email_address if first_message.sender_email_address else None

def load_pst_messages(pst, folder_name):
    folder = pst.root_folder.get_sub_folder(folder_name)
    if folder is None:
        return []
    return folder.get_contents()

def create_csv_for_pst(pst, pst_file, messages_info, source_account):
    extracts_dir = './extracts'
    os.makedirs(extracts_dir, exist_ok=True)
    csv_filename = os.path.join(extracts_dir, f"{os.path.splitext(os.path.basename(pst_file))[0]}.csv")
    total_messages = 0
    with open(csv_filename, 'w', newline='', encoding='utf-8-sig') as csvfile:
        fieldnames = ["source_account", "folder_name", "sender_email", "sender_name", "receiver_emails", "cc_emails", "bcc_emails", "delivery_time_unixtime", "subject", "attachments", "body"]
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        for folder_name, messages in messages_info.items():
            total_messages += len(messages)
            display_message_info(messages, pst, folder_name, writer, source_account)
    csv_base_name = os.path.basename(csv_filename)
    print(f"{pst_file} -> {csv_base_name} (export {total_messages} E-mails)")
    return csv_filename

def display_message_info(messages, pst, folder_name, writer, source_account):
    for message_info in messages:
        mapi_message = pst.extract_message(message_info)
        email_data = {
            "source_account": source_account,
            "folder_name": translate_folder_name(folder_name),
            "sender_email": strip_quotes(mapi_message.sender_email_address) if mapi_message.sender_email_address else '',
            "sender_name": format_kor_name(mapi_message.sender_name if mapi_message.sender_name else '').replace(" ", ""),
            "receiver_emails": strip_quotes(";".join(mapi_message.display_to.split(';') if mapi_message.display_to else ['']).strip()),
            "cc_emails": strip_quotes(mapi_message.display_cc if mapi_message.display_cc else ''),
            "bcc_emails": strip_quotes(mapi_message.display_bcc if mapi_message.display_bcc else ''),
            "delivery_time_unixtime": int(adjust_timezone(mapi_message.delivery_time, '-u9' in sys.argv).timestamp()),
            "subject": mapi_message.subject if mapi_message.subject else '',
            "attachments": ", ".join([attachment.display_name for attachment in mapi_message.attachments if attachment.display_name]),
            "body": remove_double_spaces(extract_recent_content(mapi_message.body[:2000])) if mapi_message.body else ''
        }
        writer.writerow(email_data)

def translate_folder_name(folder_name):
    folder_map = {
        "Inbox": "받은 편지함",
        "Outbox": "보낼 편지함",
        "Sent Items": "보낸 편지함",
        "Deleted Items": "삭제된 항목",
        "Drafts": "임시 보관함",
        "Junk Email": "정크 메일"
    }
    return folder_map.get(folder_name, folder_name)

def strip_quotes(text):
    return text.strip("'")

def adjust_timezone(dt, use_utc_plus_9):
    if use_utc_plus_9:
        return dt.replace(tzinfo=timezone.utc).astimezone(timezone(timedelta(hours=9)))
    return dt.replace(tzinfo=timezone.utc)

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
    
def extract_recent_content(body: str) -> str:
    lines = body.splitlines()
    result = []
    
    for i, line in enumerate(lines):
        if line[:5] == "From:":
            for j in range(1, 4):
                if i + j < len(lines):
                    sub_line = lines[i + j]
                    if sub_line[:5] == "Sent:" or sub_line[:3] == "To:" or sub_line[:8] == "Subject:":
                        return "\n".join(result)
            break
        else:
            result.append(line)
    
    return "\n".join(result)

def remove_double_spaces(body):
    while '  ' in body:
        body = body.replace('  ', ' ')
    return body

def pst_to_csv(pst_file):
    with PersonalStorage.from_file(pst_file) as pst:
        source_account = get_source_account(pst)

        folder_names = ["Inbox", "Outbox", "Sent Items", "Deleted Items", "Drafts", "Junk Email", "받은 편지함", "보낼 편지함", "보낸 편지함", "삭제된 항목", "정크 메일"]
        messages_info = {}
        for folder_name in folder_names:
            messages = load_pst_messages(pst, folder_name)
            messages_info[folder_name] = messages

        create_csv_for_pst(pst, pst_file, messages_info, source_account)

# ==================== merge_and_sort_csv_files ==================== #

def merge_and_sort_csv_files(directory):
    csv_files = glob.glob(os.path.join(directory, '*.csv'))
    all_data = []
    fieldnames = ["source_account", "folder_name", "sender_email", "sender_name", "receiver_emails", "cc_emails", "bcc_emails", "delivery_time_unixtime", "subject", "attachments", "body"]
    num_files_merged = len(csv_files)
    
    for csv_file in csv_files:
        with open(csv_file, 'r', newline='', encoding='utf-8-sig') as file:
            reader = csv.DictReader(file)
            for row in reader:
                all_data.append(row)

    all_data.sort(key=lambda x: int(x['delivery_time_unixtime']) if x['delivery_time_unixtime'] else 0)

    merged_filename = 'extract.csv'
    with open(merged_filename, 'w', newline='', encoding='utf-8-sig') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        for data in all_data:
            writer.writerow(data)

    total_rows = len(all_data)
    print(f"\n{num_files_merged} CSV files merged and sorted into '{merged_filename}' ({total_rows} E-mails)\n")

if __name__ == "__main__":
    if '-u9' in sys.argv:
        sys.argv.remove('-u9')
    
    if len(sys.argv) < 2:
        print("Usage: PST-Mail-Parser.exe [-u9] <PST file path 1> <PST file path 2> ...")
        sys.exit(1)
    
    for pst_file in sys.argv[1:]:
        pst_to_csv(pst_file)
    
    merge_and_sort_csv_files('./extracts')
