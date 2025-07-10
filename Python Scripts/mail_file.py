import win32com.client as win32  # For Outlook automation
import os  # For working with file paths
import shutil  # For moving files
import codecs  # For reading HTML email signatures




# Get the project root (dynamic365)
project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))


# Define folders
formats_path = os.path.join(project_root, "Formats")
extracted_csv_path = os.path.join(project_root, "Extracted_CSV")
mailed_csv_path = os.path.join(project_root, "Mailed_CSV")
processed_path = os.path.join(project_root, "Processed")


def __Emailer_func(folder, text, subject, recipients, p1, p2, missing=None, auto=True):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    if hasattr(recipients, 'strip'):
        recipients = [recipients]

    for recipient in recipients:
        mail.Recipients.Add(recipient)

    mail.Subject = subject
    mail.HtmlBody = text


    extract = os.path.join(extracted_csv_path, folder, p1)
    print(extract)
    mail.Attachments.Add(os.path.join(extracted_csv_path, folder, p1))
    mail.Attachments.Add(os.path.join(formats_path, folder, p2))
    
    if missing is not None:
        mail.Attachments.Add(missing)

    signature_path = os.path.join(os.environ['USERPROFILE'], 'AppData\\Roaming\\Microsoft\\Signatures\\')
    html_doc = os.path.join(signature_path, '__path__') # signature file name

    html_file = codecs.open(html_doc, 'r', 'utf-8', errors='ignore')
    signature_code = html_file.read()
    signature_code = signature_code.replace('Work_files/', signature_path)
    html_file.close()

    mail.HTMLbody = text + '\n\n' + signature_code

    if auto:
        mail.Send()
        print('Mail Sent')
    else:
        mail.Display(True)
        mail.Send()

    try:
        shutil.move(os.path.join(extracted_csv_path, folder, p1),
                    os.path.join(mailed_csv_path, folder, p1))
    except Exception as e:
        print("CSV move error:", e)

    try:
        shutil.move(os.path.join(formats_path, folder, p2),
                    os.path.join(processed_path, folder, p2))
    except Exception as e:
        print("PDF move error:", e)


def failure_mail(folder, text, subject, recipients, p1=None, auto=True):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    if hasattr(recipients, 'strip'):
        recipients = [recipients]

    for recipient in recipients:
        mail.Recipients.Add(recipient)

    mail.Subject = subject
    mail.HtmlBody = text

    if p1 is not None:
        mail.Attachments.Add(os.path.join(formats_path, folder, p1))

    signature_path = os.path.join(os.environ['USERPROFILE'], 'AppData\\Roaming\\Microsoft\\Signatures\\')
    html_doc = os.path.join(signature_path, '__path__')  # signature file name

    html_file = codecs.open(html_doc, 'r', 'utf-8', errors='ignore')
    signature_code = html_file.read()
    signature_code = signature_code.replace('Work_files/', signature_path)
    html_file.close()

    mail.HTMLbody = text + '\n\n' + signature_code

    if auto:
        mail.Send()
        print('Mail Sent')

    else:
        mail.Display(True)
        mail.Send()


