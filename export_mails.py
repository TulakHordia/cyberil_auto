import os
import win32com.client
import time
import re


def export_emails_to_keyword_folders(folder, output_folder):
    counter = 1
    pattern_ip = r'\bip\b'
    pattern_domain = r'\bdomain\b'
    pattern_url = r'\burl\b'
    pattern_md5 = r'\bmd5\b'
    pattern_sha256 = r'\bsha256\b'
    pattern_sha1 = r'\bsha1\b'
    pattern_domain_phishing = r'\b#Phishing\b'
    for message in folder.Items:
        if message.UnRead:
            body = message.Body.lower()  # Retrieves the content of the current email
            if re.search(pattern_ip, body):
                save_email_no_filter(output_folder, 'ip', message, counter)
                save_ip_tag(message)
            elif re.search(pattern_domain and pattern_domain_phishing, body):
                save_email_no_filter(output_folder, 'domain_phishing', message, counter)
                save_domain_phishing_tag(message)
            elif re.search(pattern_domain, body):
                save_email_no_filter(output_folder, 'domain', message, counter)
                save_domain_tag(message)
            # URL Keyword saved separately
            elif re.search(pattern_url, body):
                save_email_no_filter(output_folder, 'url', message, counter)
                save_url_tag(message)
            # Hash keywords
            elif re.search(pattern_md5, body):
                save_email_no_filter(output_folder, 'md5', message, counter)
                save_md5_tag(message)
            elif re.search(pattern_sha256, body):
                save_email_no_filter(output_folder, 'sha256', message, counter)
                save_sha256_tag(message)
            elif re.search(pattern_sha1, body):
                save_email_no_filter(output_folder, 'sha1', message, counter)
                save_sha1_tag(message)
            # Mark as read after tagging the message
            message.UnRead = False
            # Reset body variable for the next iteration
            body = None
            counter += 1


def save_email_no_filter(output_folder, keyword, message, counter):
    keyword_not_filtered = keyword + '_not_filtered'
    not_filtered_folder = os.path.join(output_folder, 'not_filtered')
    keyword_folder = os.path.join(not_filtered_folder, keyword_not_filtered)

    # Create the main "not_filtered" folder if it doesn't exist
    if not os.path.exists(not_filtered_folder):
        os.makedirs(not_filtered_folder)

    # Create the keyword subfolder within the "not_filtered" folder
    if not os.path.exists(keyword_folder):
        os.makedirs(keyword_folder)

    filename = f'exported_file_{keyword}_({counter}).txt'  # Modify filename to include counter
    filepath = os.path.join(keyword_folder, filename)

    with open(filepath, 'w', encoding='utf-8') as export_file:
        write_whole_email_to_txt(export_file, message)


def save_ip_tag(message):
    txt_files_folder = r'C:\Users\Kosta\Desktop\cyber.feed'

    if not os.path.exists(txt_files_folder):
        os.makedirs(txt_files_folder)

    filename = f'feedlistip.txt'  # Where to save the IPs
    filepath = os.path.join(txt_files_folder, filename)

    with open(filepath, 'a', encoding='utf-8') as export_file:
        append_subject_to_file(export_file, message)


def save_domain_tag(message):
    txt_files_folder = r'C:\Users\Kosta\Desktop\cyber.feed'

    if not os.path.exists(txt_files_folder):
        os.makedirs(txt_files_folder)

    filename = f'feedlistdomain.txt'  # Where to save the Domains
    filepath = os.path.join(txt_files_folder, filename)

    with open(filepath, 'a', encoding='utf-8') as export_file:
        append_subject_to_file(export_file, message)


def save_domain_phishing_tag(message):
    txt_files_folder = r'C:\Users\Kosta\Desktop\cyber.feed'

    if not os.path.exists(txt_files_folder):
        os.makedirs(txt_files_folder)

    filename = f'feedlistdomainPhishing.txt'  # Where to save the Domains
    filepath = os.path.join(txt_files_folder, filename)

    with open(filepath, 'a', encoding='utf-8') as export_file:
        append_subject_to_file(export_file, message)


def save_sha256_tag(message):
    txt_files_folder = r'C:\Users\Kosta\Desktop\cyber.feed'

    if not os.path.exists(txt_files_folder):
        os.makedirs(txt_files_folder)

    filename = f'feedlistsha256.txt'
    filepath = os.path.join(txt_files_folder, filename)

    with open(filepath, 'a', encoding='utf-8') as export_file:
        append_subject_to_file(export_file, message)

def save_sha1_tag(message):
    txt_files_folder = r'C:\Users\Kosta\Desktop\cyber.feed'

    if not os.path.exists(txt_files_folder):
        os.makedirs(txt_files_folder)

    filename = f'feedlistsha1.txt'
    filepath = os.path.join(txt_files_folder, filename)

    with open(filepath, 'a', encoding='utf-8') as export_file:
        append_subject_to_file(export_file, message)

def save_md5_tag(message):
    txt_files_folder = r'C:\Users\Kosta\Desktop\cyber.feed'

    if not os.path.exists(txt_files_folder):
        os.makedirs(txt_files_folder)

    filename = f'feedlistmd5.txt'
    filepath = os.path.join(txt_files_folder, filename)

    with open(filepath, 'a', encoding='utf-8') as export_file:
        append_subject_to_file(export_file, message)


def save_url_tag(message):
    txt_files_folder = r'C:\Users\Kosta\Desktop\cyber.feed'

    if not os.path.exists(txt_files_folder):
        os.makedirs(txt_files_folder)

    filename = f'feedlistURL.txt'  # Where to save the urls
    filepath = os.path.join(txt_files_folder, filename)

    with open(filepath, 'a', encoding='utf-8') as export_file:
        append_subject_to_file(export_file, message)


def append_subject_to_file(file, message):
    file.write(f"{message.Subject}\n")


def write_whole_email_to_txt(file, message):
    file.write(f"Subject: {message.Subject}\n")
    file.write(f"Received Time: {message.ReceivedTime}\n")
    file.write(f"Sender: {message.SenderName}\n")
    file.write(f"Body:\n{message.Body}\n\n")


if __name__ == "__main__":
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    
    # Iterate through top-level folders
    for folder in outlook.Folders:
    # Check if the folder is your main inbox
        if folder.Name == "kosta@itcare.co.il":
            # Access the 'cyberil.feed' folder within the main inbox
            try:
                cyberil_folder = folder.Folders['cyberil']
                # If found, you can work with the cyberil_folder here
                print("Found 'cyberil' folder!")
            except KeyError:
                print("The 'cyberil' folder was not found within the main inbox.")
            break  # Exit loop after finding the main inbox
    else:
        print("Your main inbox was not found.")
        
    output_folder = r'C:\Users\Kosta\Desktop\cyber.feed'

    export_emails_to_keyword_folders(cyberil_folder, output_folder)
    print("Exported emails to respective keyword folders.")
    time.sleep(3)  # Pause for 1 second

    # Prompt the user to press Enter to continue
    input("Press Enter to exit...")
