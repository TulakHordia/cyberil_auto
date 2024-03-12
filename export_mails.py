import os
import win32com.client

def export_emails_to_keyword_folders():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Assuming the second inbox is the index 1 in the Folders collection
    second_inbox = outlook.Folders[1].Folders['Inbox']

    specific_path = r'C:\Users\Benjamin\Desktop\Work\Test'
    ip_folder_path = os.path.join(specific_path, 'ip')
    url_folder_path = os.path.join(specific_path, 'url')
    domain_folder_path = os.path.join(specific_path, 'domain')

    # Create folders if they don't exist
    for folder_path in [ip_folder_path, url_folder_path, domain_folder_path]:
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)

    ip_counter = 1
    url_counter = 1
    domain_counter = 1

    for message in second_inbox.Items:
        if "ip" in message.Body.lower() and message.SenderName.lower() == "kosta kovshov":
            with open(os.path.join(ip_folder_path, f'exported_file({ip_counter}).txt'), 'w', encoding='utf-8') as export_file:
                export_file.write(f"Subject: {message.Subject}\n")
                export_file.write(f"Received Time: {message.ReceivedTime}\n")
                export_file.write(f"Sender: {message.SenderName}\n")
                export_file.write(f"Body:\n{message.Body}\n\n")
            ip_counter += 1
        elif "url" in message.Body.lower() and message.SenderName.lower() == "kosta kovshov":
            with open(os.path.join(url_folder_path, f'exported_file({url_counter}).txt'), 'w', encoding='utf-8') as export_file:
                export_file.write(f"Subject: {message.Subject}\n")
                export_file.write(f"Received Time: {message.ReceivedTime}\n")
                export_file.write(f"Sender: {message.SenderName}\n")
                export_file.write(f"Body:\n{message.Body}\n\n")
            url_counter += 1
        elif "domain" in message.Body.lower() and message.SenderName.lower() == "kosta kovshov":
            with open(os.path.join(domain_folder_path, f'exported_file({domain_counter}).txt'), 'w', encoding='utf-8') as export_file:
                export_file.write(f"Subject: {message.Subject}\n")
                export_file.write(f"Received Time: {message.ReceivedTime}\n")
                export_file.write(f"Sender: {message.SenderName}\n")
                export_file.write(f"Body:\n{message.Body}\n\n")
            domain_counter += 1

    print("Exported emails to respective keyword folders.")

if __name__ == "__main__":
    export_emails_to_keyword_folders()
