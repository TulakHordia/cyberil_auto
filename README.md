# Outlook Email Exporter

This Python script extracts emails from a specified Outlook folder and exports them to keyword-specific folders based on the content of the emails. It utilizes the `win32com.client` library to interact with Outlook.

## Features

- Extracts emails containing specific keywords (e.g., "IP", "URL", "Domain", "MD5", "SHA256").
- Saves emails to keyword-specific folders.
- Extracts and saves subject lines of emails containing specific keywords to separate text files.

## Requirements

- Python 3.x
- `win32com` library (install via `pip install pywin32`)

## Usage

1. Make sure you have Python installed on your system.
2. Install the required `win32com` library by running `pip install pywin32`.
3. Update the script with your desired input and output folder paths.
4. Run the script.

## Instructions

1. Modify the `output_folder` variable in the script to specify the directory where you want the emails to be exported.
2. Ensure that Microsoft Outlook is installed on your system and configured with the desired email account.
3. Modify the `cyberil_folder` variable in the script to specify the Outlook folder from which you want to extract emails.
4. Run the script. Extracted emails will be saved in keyword-specific folders within the specified output directory.

## Notes

- This script currently supports extraction and export of emails containing keywords "IP", "URL", "Domain", "MD5", and "SHA256". You can extend it to include more keywords as needed.
- Make sure to review and adapt the script according to your specific requirements and file organization preferences.

## Disclaimer

- Use this script responsibly and in compliance with applicable laws and regulations.
- The script author assumes no liability for any misuse or damage caused by the use of this script.
