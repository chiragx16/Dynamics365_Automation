import win32com.client  # For Outlook automation
import re  # For regex-based filename cleanup
import os  # For file handling
import time


def ensure_directory_exists(path):
    """Ensure that the target directory exists; if not, create it."""
    if not os.path.exists(path):
        os.makedirs(path)


def Download():
    time.sleep(5)  # Delay to ensure Outlook is ready
    print("Downloading Mails...")
    # Get the correct "Formats" folder inside "dynamic365"
    base_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "Formats"))

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # Inbox folder
    messages = inbox.Items

    for message in messages:
        if message.UnRead:  # Process only unread messages
            subject = message.Subject.lower()  # Convert subject to lowercase for case-insensitive search
            
            if "dynamics365" in subject:  
                print("Matching Email Found:", subject)
                
                attachments = message.Attachments

                for attachment in attachments:
                    if attachment.FileName.endswith(('pdf', 'PDF')):

                        # Preserve spaces while removing special characters
                        filename = re.sub(r'[^0-9a-zA-Z\s\.\-]+', '', attachment.FileName)

                        # Define subfolder based on subject keywords
                        subfolder = None
                        if "dynamics365 military" in subject:
                            subfolder = "MIL"
                        elif "dynamics365 reliance" in subject:
                            subfolder = "RIL"
                        elif "dynamics365 dmart" in subject:
                            subfolder = "Dmart"
                        elif "dynamics365 csd" in subject:
                            subfolder = "CSD"
                        elif "dynamics365 aadhar" in subject:
                            subfolder = "Aadhar"
                        elif "dynamics365 lulu" in subject:
                            subfolder = "LULU"

                        if subfolder:
                            save_path = os.path.join(base_path, subfolder)
                            ensure_directory_exists(save_path)  # Ensure subfolder exists
                            attachment.SaveAsFile(os.path.join(save_path, filename))
                            print(f"Saved: {filename} → {save_path}")

                message.UnRead = False  # Mark email as read

