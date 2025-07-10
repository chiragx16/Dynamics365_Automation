import sys
import os
import pandas as pd
import time
import sys
import shutil
from sql_file import sql_formatter
from download_file import Download
from mail_file import __Emailer_func, failure_mail

from dynamics_login import dynamics_login
from payment_auto import dynamics_automation
from tds_auto import dynamics_automation_tds


from dmart_extract import extract_dmart
from lulu_extract import extract_lulu
from aadhar_extract import extract_aadhar_format
from ril_extract import extract_text_from_pdf_ril

import ast

import configparser

# Load the configuration
config = configparser.ConfigParser()
config.read('__path__') # config file path


# Read email and password from the [QA] section
recipients_str  = config.get('Prod', 'recipients')  
recipients = ast.literal_eval(recipients_str)
main_path = config.get('Prod', 'main_path')
output_path = config.get('Prod', 'output_path')


# Function to clean DF before TDS Automation
# (Removes Entries where TDS is 0 or 0.00)
def clean_dataframe(df):
    rows_to_remove = []
    df = df.reset_index(drop=True)


    for i in range(len(df)):
        tds_values = df.iloc[i, df.columns.get_loc("TDS")]

        # If TDS contains only 0.00, mark the row for removal
        if all(tds == 0.00 for tds in tds_values):
            rows_to_remove.append(i)
        else:
            # Remove corresponding values where TDS is 0.00
            indices_to_remove = [j for j, tds in enumerate(tds_values) if tds == 0.00]
            for col in ["Invoice Amount", "Payment Amount", "TDS", "Amount of Deduction", "Invoice No."]:
                df.at[i, col] = [v for j, v in enumerate(df.at[i, col]) if j not in indices_to_remove]

    # Drop rows that should be removed
    df = df.drop(index=rows_to_remove).reset_index(drop=True)
    return df

print("Starting Program 1")



# ------------------------------- Run the Program -----------------------------------
# -----------------------------------------------------------------------------------



# ------------------------------- "Download Mails" ----------------------------------

Download()


# ------------------------------- "Login to Dynamics365" -----------------------------

# email_recipients = recipients
# try:
#     driver = dynamics_login()
# except Exception as e:
#     folder = ""
#     failure_subject = f"Failure in Dynamics365 Automation - {folder}"
#     failure_body = f'<br><br>Failed to login | Dynamics365<br>'
#     file = None
#     print("Error : " ,e)
#     failure_mail(folder, failure_body, failure_subject, email_recipients , file)
#     sys.exit(1)




# email_recipients = recipients
# max_retries = 3
# attempt = 0
# driver = None

# while attempt < max_retries:
#     try:
#         driver = dynamics_login()
#         break  # Exit loop on successful login
#     except Exception as e:
#         attempt += 1
#         print(f"Login attempt {attempt} failed. Error: {e}")
        
#         if attempt == max_retries:
#             folder = ""
#             failure_subject = f"Failure in Dynamics365 Automation - {folder}"
#             failure_body = f'<br><br>Failed to login | Dynamics365 after {max_retries} attempts<br>'
#             file = None
#             failure_mail(folder, failure_body, failure_subject, email_recipients, file)
#             sys.exit(1)
#         else:
#             print("Retrying in 5 minutes...")
#             time.sleep(300)  # Wait 5 minutes before the next attempt




email_recipients = recipients
max_retries = 3
attempt = 0
driver = None

# === Folder Paths ===
base_folder = r"E:\Dynamic_365\Formats"
folders_to_check = ["Aadhar", "RIL", "LULU", "Dmart"]
full_paths = [os.path.join(base_folder, folder) for folder in folders_to_check]

# === Function to check if all folders are empty ===
def all_folders_empty(paths):
    for path in paths:
        if os.path.exists(path) and os.listdir(path):  # Not empty
            return False
    return True

# === Skip login if all folders are empty ===
if all_folders_empty(full_paths):
    print("All target folders are empty. Skipping login.")
    sys.exit(0)

print("Starting Program 2")

# === Attempt login ===
while attempt < max_retries:
    try:
        driver = dynamics_login()
        break  # Exit loop on successful login
    except Exception as e:
        attempt += 1
        print(f"Login attempt {attempt} failed. Error: {e}")
        
        if attempt == max_retries:
            folder = ""
            failure_subject = f"Failure in Dynamics365 Automation - {folder}"
            failure_body = f'<br><br>Failed to login | Dynamics365 after {max_retries} attempts<br>'
            file = None
            failure_mail(folder, failure_body, failure_subject, email_recipients, file)
            sys.exit(1)
        else:
            print("Retrying in 5 minutes...")
            time.sleep(300)  # Wait 5 minutes before next attempt


# ------------------------------- "Main Part" -----------------------------

main_dir_path = main_path
output_dir = output_path


email_recipients = recipients


for folder in os.listdir(main_dir_path):
    folder_path = os.path.join(main_dir_path, folder)
    
    if os.path.isdir(folder_path):
        print("Folder:", folder)
        for file in os.listdir(folder_path):
            if file.lower().endswith(".pdf"):
                file_path = os.path.join(folder_path, file)


                if folder == "Dmart":
                    print('-' * 150)
                    print('-' * 150)
                    print(f"Processing: {file} in {folder}")

                    # Extract PDF
                    try:
                        dic, deduction_df = extract_dmart(file_path)
                    except Exception as e:
                        failure_subject = f"Failure in Dynamics365 Automation - {folder}"
                        failure_body = f'<br><br>Error in Extracting Data From PDF<br>'
                        failure_mail(folder, failure_body, failure_subject, email_recipients, file)

                        # Move file to Failure_Files
                        try:
                            failure_dir = r"E:\Dynamic_365\Failure_Files\Dmart"
                            os.makedirs(failure_dir, exist_ok=True)  # Ensure folder exists
                            shutil.move(file_path, os.path.join(failure_dir, os.path.basename(file_path)))
                        except Exception as move_error:
                            print(f"Failed to move file: {move_error}")
                        
                        continue  

                    # Fetch from Database
                    try:
                        dic, missing_invoices_df = sql_formatter(dic)
                    except Exception as e:
                        print(e)
                        failure_subject = f"Failure in Dynamics365 Automation - {folder}"
                        failure_body = f'<br><br>Error in Fetching Data From Database<br>'
                        failure_mail(folder, failure_body, failure_subject, email_recipients , file)
                        continue  
                        
                    
                    dgh = dic[file]
                    # dgh.to_excel("1.xlsx")
                    # print(dgh)

                    # Get DataFrame from Extraction if empty send failure mail
                    if dgh is None or dgh.empty:   
                        failure_subject = f"Failure in Dynamics365 Automation - {folder}"
                        fail_body = "No invoice from pdf found in Database"
                        failure_body = f'<br><br>{fail_body}<br>'
                        failure_mail(folder, failure_body, failure_subject, email_recipients , file)
                        continue  

                        
                    print("Before Removing \n")
                    print(dgh)


                    # Perform Payment Amount Automation
                    try:    
                        driver, final_, missing_pay, mismatch_df = dynamics_automation(driver, dgh, "AVENUE", "CTR000012")
                    except Exception as e:
                        print("Error :" , e)
                        failure_subject = f"Failure in Dynamics365 Automation - {folder}"
                        failure_body = f'<br><br>Error in Performing Payment Amount Automation<br>'
                        failure_mail(folder, failure_body, failure_subject, email_recipients , file)
                        continue  
                                        

                    tds_df = clean_dataframe(final_)
                    
                    print("After Removing \n")
                    print(tds_df)


                    # Perform TDS Amount Automation
                    if tds_df is not None or not tds_df.empty:
                        try:
                            driver, final_tds, missing_tds, mismatch_tds = dynamics_automation_tds(driver, tds_df, "AVENUE", "CTR000012")
                        except Exception as e:
                            failure_subject = f"Failure in Dynamics365 Automation - {folder}"
                            failure_body = f'<br><br>Error in Performing TDS Amount Automation<br>'
                            failure_mail(folder, failure_body, failure_subject, email_recipients , file)
                            continue  
                        
                    df_exploded2 = final_tds.explode(["Invoice Amount", "Payment Amount", "Amount of Deduction", "Invoice No.", "TDS"]).reset_index(drop=True)

                    # Create Excel file
                    output_folder = os.path.join(output_dir, folder)
                    os.makedirs(output_folder, exist_ok=True)
    
                    output_file = os.path.join(output_folder, f"{file}.xlsx")
    
                    df_exploded = final_tds.explode(["Invoice Amount", "Payment Amount", "Amount of Deduction", "Invoice No.", "TDS"]).reset_index(drop=True)
                    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                        df_exploded.to_excel(writer, sheet_name='Sheet1', index=False)
                        if deduction_df is not None:
                            deduction_df.to_excel(writer, sheet_name='Sheet2', index=False)  # Another sheet


                    # Final mail Info.
                    email_subject = f"Processed - {folder}"
                    email_body = f'<br><br>{df_exploded2.to_html()}<br>'

                    # Attach Description Table if Exists
                    if deduction_df is not None:
                        email_body = email_body + f"<br><br>{deduction_df.to_html()}<br>"

                    
                    email_recipients = recipients
                    email_attachment = f"{file}.xlsx"

                    
                    missing_output_folder = os.path.join(output_dir, f"REPORT_{folder}")
                    os.makedirs(missing_output_folder, exist_ok=True)
                    missing_output_file = os.path.join(missing_output_folder, f"report_{file}.xlsx")

                    if any([
                        missing_invoices_df is not None and not missing_invoices_df.empty,
                        missing_pay is not None and not missing_pay.empty,
                        missing_tds is not None and not missing_tds.empty,
                        mismatch_df is not None and not mismatch_df.empty,
                        mismatch_tds is not None and not mismatch_tds.empty
                    ]):
                        with pd.ExcelWriter(missing_output_file, engine='xlsxwriter') as writer:
                            if missing_invoices_df is not None and not missing_invoices_df.empty:
                                missing_invoices_df.rename(columns={"Invoice No.": "Could not found Invoices in Database"}, inplace=True)
                                missing_invoices_df.to_excel(writer, sheet_name='Sheet1', index=False)
                            if missing_pay is not None and not missing_pay.empty:
                                missing_pay.to_excel(writer, sheet_name='Sheet2', index=False)
                            if missing_tds is not None and not missing_tds.empty:
                                missing_tds.to_excel(writer, sheet_name='Sheet3', index=False)
                            if mismatch_df is not None and not mismatch_df.empty:
                                mismatch_df.to_excel(writer, sheet_name='Sheet4', index=False)
                            if mismatch_tds is not None and not mismatch_tds.empty:
                                mismatch_tds.to_excel(writer, sheet_name='Sheet5', index=False)
                    
                        missing_excel = missing_output_file  # Attach the file if created
                    else:
                        missing_excel = None  # No file will be created or attached           
                        
                    __Emailer_func(folder, email_body, email_subject, email_recipients, email_attachment, file, missing_excel, auto=True)



                if folder == "RIL":
                    print('-' * 150)
                    print('-' * 150)
                    print(f"Processing: {file} in {folder}")

                    # Extract PDF
                    try:
                        dic, deduction_df = extract_text_from_pdf_ril(file_path)
                    except Exception as e:
                        failure_subject = f"Failure in Dynamics365 Automation - {folder}"
                        failure_body = f'<br><br>Error in Extracting Data From PDF<br>'
                        failure_mail(folder, failure_body, failure_subject, email_recipients, file)

                        # Move file to Failure_Files
                        try:
                            failure_dir = r"E:\Dynamic_365\Failure_Files\RIL"
                            os.makedirs(failure_dir, exist_ok=True)  # Ensure folder exists
                            shutil.move(file_path, os.path.join(failure_dir, os.path.basename(file_path)))
                        except Exception as move_error:
                            print(f"Failed to move file: {move_error}")
                        
                        continue  

                    # Fetch from Database
                    try:
                        dic, missing_invoices_df = sql_formatter(dic)
                    except Exception as e:
                        failure_subject = f"Failure in Dynamics365 Automation - {folder}"
                        failure_body = f'<br><br>Error in Fetching Data From Database<br>'
                        failure_mail(folder, failure_body, failure_subject, email_recipients , file)
                        continue  
                        
                    
                    dgh = dic[file]
                    # dgh.to_excel("dgh.xlsx")
                    # print(dgh)

                    # Get DataFrame from Extraction if empty send failure mail
                    if dgh is None or dgh.empty:   
                        failure_subject = f"Failure in Dynamics365 Automation - {folder}"
                        fail_body = "No invoice from pdf found in Database"
                        failure_body = f'<br><br>{fail_body}<br>'
                        failure_mail(folder, failure_body, failure_subject, email_recipients , file)
                        continue  

                        
                    print("Before Removing \n")
                    print(dgh)


                    # Perform Payment Amount Automation
                    try:    
                        driver, final_, missing_pay, mismatch_df = dynamics_automation(driver, dgh, "RELIANCE", "CTR000011")
                    except Exception as e:
                        print("Error :" , e)
                        failure_subject = f"Failure in Dynamics365 Automation - {folder}"
                        failure_body = f'<br><br>Error in Performing Payment Amount Automation<br>'
                        failure_mail(folder, failure_body, failure_subject, email_recipients , file)
                        continue  
                                        
                    final_.to_excel("final_.xlsx")
                    tds_df = clean_dataframe(final_)
                    
                    print("After Removing \n")
                    print(tds_df)
                    tds_df.to_excel("tds_df.xlsx")


                    # Perform TDS Amount Automation
                    if tds_df is not None or not tds_df.empty:
                        try:
                            driver, final_tds, missing_tds, mismatch_tds = dynamics_automation_tds(driver, tds_df, "RELIANCE", "CTR000011")
                        except Exception as e:
                            failure_subject = f"Failure in Dynamics365 Automation - {folder}"
                            failure_body = f'<br><br>Error in Performing TDS Amount Automation<br>'
                            failure_mail(folder, failure_body, failure_subject, email_recipients , file)
                            continue  
                        
                    df_exploded2 = final_tds.explode(["Invoice Amount", "Payment Amount", "Amount of Deduction", "Invoice No.", "TDS"]).reset_index(drop=True)

                    # Create Excel file
                    output_folder = os.path.join(output_dir, folder)
                    os.makedirs(output_folder, exist_ok=True)
    
                    output_file = os.path.join(output_folder, f"{file}.xlsx")
    
                    df_exploded = final_tds.explode(["Invoice Amount", "Payment Amount", "Amount of Deduction", "Invoice No.", "TDS"]).reset_index(drop=True)
                    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                        df_exploded.to_excel(writer, sheet_name='Sheet1', index=False)
                        if deduction_df is not None:
                            deduction_df.to_excel(writer, sheet_name='Sheet2', index=False)  # Another sheet


                    # Final mail Info.
                    email_subject = f"Processed - {folder}"
                    email_body = f'<br><br>{df_exploded2.to_html()}<br>'

                    # Attach Description Table if Exists
                    if deduction_df is not None:
                        email_body = email_body + f"<br><br>{deduction_df.to_html()}<br>"

                    
                    email_recipients = recipients
                    email_attachment = f"{file}.xlsx"

                    
                    missing_output_folder = os.path.join(output_dir, f"REPORT_{folder}")
                    os.makedirs(missing_output_folder, exist_ok=True)
                    missing_output_file = os.path.join(missing_output_folder, f"report_{file}.xlsx")

                    if any([
                        missing_invoices_df is not None and not missing_invoices_df.empty,
                        missing_pay is not None and not missing_pay.empty,
                        missing_tds is not None and not missing_tds.empty,
                        mismatch_df is not None and not mismatch_df.empty,
                        mismatch_tds is not None and not mismatch_tds.empty
                    ]):
                        with pd.ExcelWriter(missing_output_file, engine='xlsxwriter') as writer:
                            if missing_invoices_df is not None and not missing_invoices_df.empty:
                                missing_invoices_df.rename(columns={"Invoice No.": "Could not found Invoices in Database"}, inplace=True)
                                missing_invoices_df.to_excel(writer, sheet_name='Sheet1', index=False)
                            if missing_pay is not None and not missing_pay.empty:
                                missing_pay.to_excel(writer, sheet_name='Sheet2', index=False)
                            if missing_tds is not None and not missing_tds.empty:
                                missing_tds.to_excel(writer, sheet_name='Sheet3', index=False)
                            if mismatch_df is not None and not mismatch_df.empty:
                                mismatch_df.to_excel(writer, sheet_name='Sheet4', index=False)
                            if mismatch_tds is not None and not mismatch_tds.empty:
                                mismatch_tds.to_excel(writer, sheet_name='Sheet5', index=False)
                    
                        missing_excel = missing_output_file  # Attach the file if created
                    else:
                        missing_excel = None  # No file will be created or attached           
                        
                    __Emailer_func(folder, email_body, email_subject, email_recipients, email_attachment, file, missing_excel, auto=True)
    


                if folder == "Aadhar":
                    print('-' * 150)
                    print('-' * 150)
                    print(f"Processing: {file} in {folder}")

                    # Extract PDF
                    try:
                        dic, deduction_df = extract_aadhar_format(file_path)
                    except Exception as e:
                        failure_subject = f"Failure in Dynamics365 Automation - {folder}"
                        failure_body = f'<br><br>Error in Extracting Data From PDF<br>'
                        failure_mail(folder, failure_body, failure_subject, email_recipients, file)

                        # Move file to Failure_Files
                        try:
                            failure_dir = r"E:\Dynamic_365\Failure_Files\Aadhar"
                            os.makedirs(failure_dir, exist_ok=True)  # Ensure folder exists
                            shutil.move(file_path, os.path.join(failure_dir, os.path.basename(file_path)))
                        except Exception as move_error:
                            print(f"Failed to move file: {move_error}")
                        
                        continue  

                    # Fetch from Database
                    try:
                        dic, missing_invoices_df = sql_formatter(dic)
                    except Exception as e:
                        failure_subject = f"Failure in Dynamics365 Automation - {folder}"
                        failure_body = f'<br><br>Error in Fetching Data From Database<br>'
                        failure_mail(folder, failure_body, failure_subject, email_recipients , file)
                        continue  
                        
                    
                    dgh = dic[file]
                    # dgh.to_excel("dgh.xlsx")
                    # print(dgh)

                    # Get DataFrame from Extraction if empty send failure mail
                    if dgh is None or dgh.empty:   
                        failure_subject = f"Failure in Dynamics365 Automation - {folder}"
                        fail_body = "No invoice from pdf found in Database"
                        failure_body = f'<br><br>{fail_body}<br>'
                        failure_mail(folder, failure_body, failure_subject, email_recipients , file)
                        continue  

                        
                    print("Before Removing \n")
                    print(dgh)


                    # Perform Payment Amount Automation
                    try:    
                        driver, final_, missing_pay, mismatch_df = dynamics_automation(driver, dgh, "AADHAR", "CTR000027")
                    except Exception as e:
                        print("Error :" , e)
                        failure_subject = f"Failure in Dynamics365 Automation - {folder}"
                        failure_body = f'<br><br>Error in Performing Payment Amount Automation<br>'
                        failure_mail(folder, failure_body, failure_subject, email_recipients , file)
                        continue  
                                        

                    tds_df = clean_dataframe(final_)
                    
                    print("After Removing \n")
                    print(tds_df)


                    # Perform TDS Amount Automation
                    if tds_df is not None or not tds_df.empty:
                        try:
                            driver, final_tds, missing_tds, mismatch_tds = dynamics_automation_tds(driver, tds_df, "AADHAR", "CTR000027")
                        except Exception as e:
                            failure_subject = f"Failure in Dynamics365 Automation - {folder}"
                            failure_body = f'<br><br>Error in Performing TDS Amount Automation<br>'
                            failure_mail(folder, failure_body, failure_subject, email_recipients , file)
                            continue  
                        
                    df_exploded2 = final_tds.explode(["Invoice Amount", "Payment Amount", "Amount of Deduction", "Invoice No.", "TDS"]).reset_index(drop=True)

                    # Create Excel file
                    output_folder = os.path.join(output_dir, folder)
                    os.makedirs(output_folder, exist_ok=True)
    
                    output_file = os.path.join(output_folder, f"{file}.xlsx")
    
                    df_exploded = final_tds.explode(["Invoice Amount", "Payment Amount", "Amount of Deduction", "Invoice No.", "TDS"]).reset_index(drop=True)
                    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                        df_exploded.to_excel(writer, sheet_name='Sheet1', index=False)
                        if deduction_df is not None:
                            deduction_df.to_excel(writer, sheet_name='Sheet2', index=False)  # Another sheet


                    # Final mail Info.
                    email_subject = f"Processed - {folder}"
                    email_body = f'<br><br>{df_exploded2.to_html()}<br>'

                    # Attach Description Table if Exists
                    if deduction_df is not None:
                        email_body = email_body + f"<br><br>{deduction_df.to_html()}<br>"

                    
                    email_recipients = recipients
                    email_attachment = f"{file}.xlsx"

                    
                    missing_output_folder = os.path.join(output_dir, f"REPORT_{folder}")
                    os.makedirs(missing_output_folder, exist_ok=True)
                    missing_output_file = os.path.join(missing_output_folder, f"report_{file}.xlsx")

                    if any([
                        missing_invoices_df is not None and not missing_invoices_df.empty,
                        missing_pay is not None and not missing_pay.empty,
                        missing_tds is not None and not missing_tds.empty,
                        mismatch_df is not None and not mismatch_df.empty,
                        mismatch_tds is not None and not mismatch_tds.empty
                    ]):
                        with pd.ExcelWriter(missing_output_file, engine='xlsxwriter') as writer:
                            if missing_invoices_df is not None and not missing_invoices_df.empty:
                                missing_invoices_df.rename(columns={"Invoice No.": "Could not found Invoices in Database"}, inplace=True)
                                missing_invoices_df.to_excel(writer, sheet_name='Sheet1', index=False)
                            if missing_pay is not None and not missing_pay.empty:
                                missing_pay.to_excel(writer, sheet_name='Sheet2', index=False)
                            if missing_tds is not None and not missing_tds.empty:
                                missing_tds.to_excel(writer, sheet_name='Sheet3', index=False)
                            if mismatch_df is not None and not mismatch_df.empty:
                                mismatch_df.to_excel(writer, sheet_name='Sheet4', index=False)
                            if mismatch_tds is not None and not mismatch_tds.empty:
                                mismatch_tds.to_excel(writer, sheet_name='Sheet5', index=False)
                    
                        missing_excel = missing_output_file  # Attach the file if created
                    else:
                        missing_excel = None  # No file will be created or attached           
                        
                    __Emailer_func(folder, email_body, email_subject, email_recipients, email_attachment, file, missing_excel, auto=True)
    

                
                if folder == "LULU":
                    print('-' * 150)
                    print('-' * 150)
                    print(f"Processing: {file} in {folder}")

                    # Extract PDF
                    try:
                        dic, deduction_df = extract_lulu(file_path)
                    except Exception as e:
                        failure_subject = f"Failure in Dynamics365 Automation - {folder}"
                        failure_body = f'<br><br>Error in Extracting Data From PDF<br>'
                        failure_mail(folder, failure_body, failure_subject, email_recipients, file)

                        # Move file to Failure_Files
                        try:
                            failure_dir = r"E:\Dynamic_365\Failure_Files\LULU"
                            os.makedirs(failure_dir, exist_ok=True)  # Ensure folder exists
                            shutil.move(file_path, os.path.join(failure_dir, os.path.basename(file_path)))
                        except Exception as move_error:
                            print(f"Failed to move file: {move_error}")
                        
                        continue  

                    # Fetch from Database
                    try:
                        dic, missing_invoices_df = sql_formatter(dic)
                    except Exception as e:
                        failure_subject = f"Failure in Dynamics365 Automation - {folder}"
                        failure_body = f'<br><br>Error in Fetching Data From Database<br>'
                        failure_mail(folder, failure_body, failure_subject, email_recipients , file)
                        continue  
                        
                    
                    dgh = dic[file]
                    # dgh.to_excel("dgh.xlsx")
                    # print(dgh)

                    # Get DataFrame from Extraction if empty send failure mail
                    if dgh is None or dgh.empty:   
                        failure_subject = f"Failure in Dynamics365 Automation - {folder}"
                        fail_body = "No invoice from pdf found in Database"
                        failure_body = f'<br><br>{fail_body}<br>'
                        failure_mail(folder, failure_body, failure_subject, email_recipients , file)
                        continue  

                        
                    print("Before Removing \n")
                    print(dgh)


                    # Perform Payment Amount Automation
                    try:    
                        driver, final_, missing_pay, mismatch_df = dynamics_automation(driver, dgh, "LULU", "CTR000025")
                    except Exception as e:
                        print("Error :" , e)
                        failure_subject = f"Failure in Dynamics365 Automation - {folder}"
                        failure_body = f'<br><br>Error in Performing Payment Amount Automation<br>'
                        failure_mail(folder, failure_body, failure_subject, email_recipients , file)
                        continue  
                                        

                        
                    df_exploded2 = final_.explode(["Invoice Amount", "Payment Amount", "Amount of Deduction", "Invoice No.", "TDS"]).reset_index(drop=True)

                    # Create Excel file
                    output_folder = os.path.join(output_dir, folder)
                    os.makedirs(output_folder, exist_ok=True)
    
                    output_file = os.path.join(output_folder, f"{file}.xlsx")
    
                    df_exploded = final_.explode(["Invoice Amount", "Payment Amount", "Amount of Deduction", "Invoice No.", "TDS"]).reset_index(drop=True)
                    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                        df_exploded.to_excel(writer, sheet_name='Sheet1', index=False)
                        if deduction_df is not None:
                            deduction_df.to_excel(writer, sheet_name='Sheet2', index=False)  # Another sheet


                    # Final mail Info.
                    email_subject = f"Processed - {folder}"
                    email_body = f'<br><br>{df_exploded2.to_html()}<br>'

                    # Attach Description Table if Exists
                    if deduction_df is not None:
                        email_body = email_body + f"<br><br>{deduction_df.to_html()}<br>"

                    
                    email_recipients = recipients
                    email_attachment = f"{file}.xlsx"

                    
                    missing_output_folder = os.path.join(output_dir, f"REPORT_{folder}")
                    os.makedirs(missing_output_folder, exist_ok=True)
                    missing_output_file = os.path.join(missing_output_folder, f"report_{file}.xlsx")

                    if any([
                        missing_invoices_df is not None and not missing_invoices_df.empty,
                        missing_pay is not None and not missing_pay.empty,
                        missing_tds is not None and not missing_tds.empty,
                        mismatch_df is not None and not missing_tds.empty
                    ]):
                        with pd.ExcelWriter(missing_output_file, engine='xlsxwriter') as writer:
                            if missing_invoices_df is not None and not missing_invoices_df.empty:
                                missing_invoices_df.rename(columns={"Invoice No.": "Could not found Invoices in Database"}, inplace=True)
                                missing_invoices_df.to_excel(writer, sheet_name='Sheet1', index=False)
                            if missing_pay is not None and not missing_pay.empty:
                                missing_pay.to_excel(writer, sheet_name='Sheet2', index=False)                  
                            if mismatch_df is not None and not mismatch_df.empty:
                                mismatch_df.to_excel(writer, sheet_name='Sheet3', index=False)

                    
                        missing_excel = missing_output_file  # Attach the file if created
                    else:
                        missing_excel = None  # No file will be created or attached           
                        
                    __Emailer_func(folder, email_body, email_subject, email_recipients, email_attachment, file, missing_excel, auto=True)
    

