import os
import win32com.client


def export_emails_to_keyword_folders(folder, output_folder):
    counter = 1
    for message in folder.Items:
        body = message.Body.lower()
        if "ip" in body:
            save_email_no_filter(output_folder, 'ip', message, counter)
            save_email_and_filter(output_folder, 'ip', message)
            save_ip_tag(output_folder, 'ip', message)
        elif "url" in body:
            save_email_no_filter(output_folder, 'url', message, counter)
            save_email_and_filter(output_folder, 'url', message)
            save_url_tag(output_folder, 'url', message)
        elif "domain" in body:
            save_email_no_filter(output_folder, 'domain', message, counter)
            save_email_and_filter(output_folder, 'domain', message)
            save_domain_tag(output_folder, 'domain', message)
        elif "md5" in body:
            save_email_no_filter(output_folder, 'md5', message, counter)
            save_email_and_filter(output_folder, 'md5', message)
        elif "sha256" in body:
            save_email_no_filter(output_folder, 'sha256', message, counter)
            save_email_and_filter(output_folder, 'sha256', message)
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

def save_ip_tag(output_folder, keyword, message):
    txt_files_folder = r'C:\Users\Something'
    filename = f'exported_file_{keyword}.txt'  # Modify filename to include counter
    filepath = os.path.join(txt_files_folder, filename)

    with open(filepath, 'a', encoding='utf-8') as export_file:
        append_subject_to_file(export_file, message)
def save_domain_tag(output_folder, keyword, message):
    txt_files_folder = r'C:\Users\Something'
    filename = f'exported_file_{keyword}.txt'  # Modify filename to include counter
    filepath = os.path.join(txt_files_folder, filename)

    with open(filepath, 'a', encoding='utf-8') as export_file:
        append_subject_to_file(export_file, message)
def save_url_tag(output_folder, keyword, message):
    txt_files_folder = r'C:\Users\Something'
    filename = f'exported_file_{keyword}.txt'  # Modify filename to include counter
    filepath = os.path.join(txt_files_folder, filename)

    with open(filepath, 'a', encoding='utf-8') as export_file:
        append_subject_to_file(export_file, message)


def save_email_and_filter(output_folder, keyword, message):
    keyword_with_filtered = keyword + '_filtered'
    filtered_folder = os.path.join(output_folder, 'filtered')
    keyword_folder = os.path.join(filtered_folder, keyword_with_filtered)

    # Create the main "not_filtered" folder if it doesn't exist
    if not os.path.exists(filtered_folder):
        os.makedirs(filtered_folder)

    # Create the keyword subfolder within the "not_filtered" folder
    if not os.path.exists(keyword_folder):
        os.makedirs(keyword_folder)

    filename = f'exported_file_{keyword}.txt'  # Modify filename to include counter
    filepath = os.path.join(keyword_folder, filename)

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
    cyberil_folder = outlook.Folders[1].Folders['cyberil']
    output_folder = r'C:\Users\geshe\Dropbox\Python\Test'
    export_emails_to_keyword_folders(cyberil_folder, output_folder)
    print("Exported emails to respective keyword folders.")

