import os
import re
import pandas as pd
import pdfplumber



def extract_pay_advise_no(text):
    matches = re.findall(r"Payment Ref\. No\. ?: ?([^\n]+)", text)
    return matches

def extract_table_data(pdf_path):
    beneficiary_code = None
    extracted_data = []
    
    with pdfplumber.open(pdf_path) as pdf: # Loop through pages
        for page in pdf.pages:
            text = page.extract_text()

            # print(repr(text))
            if text:
                if not beneficiary_code:
                    beneficiary_code = extract_pay_advise_no(text)
                
                table_start_match = re.search(r"Sr\.No\.\s+Invoice Number\s+Invoice Date", text)
                if table_start_match:
                    table_start_index = table_start_match.end()
                    table_data = text[table_start_index:].strip()
                    
                    lines = table_data.split("\n")
                    i = 0
                    while i < len(lines):
                        if "*The time taken for effective credit" in lines[i]:
                            break
                        
                        main_match = re.match(r'(\d+)\s+(.+?)\s+([\d-]+)\s+([\d.]+)\s+([\d.]+)\s+([-?\d,.\d]+)\s+(\d+)\s+(.+)', lines[i])
                        
                        if main_match:
                            invoice_no = main_match.group(2).strip()
                            
                            if not re.match(r'^[A-Za-z0-9-\s]+$', invoice_no):
                                j = i + 1
                                while j < len(lines):
                                    if not re.match(r'(\d+)\s+(.+?)\s+([\d-]+)\s+([\d.]+)\s+([\d.]+)\s+([-?\d,.\d]+)\s+(\d+)\s+(.+)', lines[j]):
                                        invoice_no += " " + lines[j].strip()
                                        j += 1
                                    else:
                                        break
                                i = j - 1
                            
                            extracted_data.append({
                                "Sr.No.": main_match.group(1),
                                "Invoice No.": invoice_no,
                                "Invoice Date": main_match.group(3),
                                "TCS Collected": main_match.group(4),
                                "TDS": main_match.group(5),
                                "Payment Amount": main_match.group(6),
                                "Site": main_match.group(7),
                                "Site Name": main_match.group(8),
                            })
                        
                        i += 1
    
    return beneficiary_code, extracted_data

def extract_dmart(pdf_path):
    pdf_filename = os.path.basename(pdf_path)
    final_dc = {}
    
    beneficiary_code, table_data = extract_table_data(pdf_path)
    
    final_df = pd.DataFrame(table_data)
    dmart_deductions = pd.DataFrame()
    
    final_df["Payment Amount"] = pd.to_numeric(final_df["Payment Amount"].str.replace(",", ""), errors="coerce")
    final_df["TDS"] = pd.to_numeric(final_df["TDS"].str.replace(",", ""), errors="coerce")
    
    final_df["Invoice Amount"] = final_df["Payment Amount"] + final_df["TDS"]
    final_df["Amount of Deduction"] = final_df["Invoice Amount"] - final_df["Payment Amount"]
    
    dmart_deductions = final_df[final_df["Payment Amount"] < 0]
    final_df = final_df[final_df["Payment Amount"] >= 0]
    
    # final_df["Other Amount"] = 0.0
    final_df["Payment Advice No."] = beneficiary_code[0]

    if not final_df.empty:
        final_df.reset_index(drop=True, inplace=True)
    
    # Remove any entries in 'Invoice No.' containing '*The time'
    dmart_deductions['Invoice No.'] = dmart_deductions['Invoice No.'].str.replace(r'The time.*', '', regex=True)
    dmart_deductions["Payment Advice No."] = beneficiary_code[0]

    if dmart_deductions is not None or not dmart_deductions.empty:
            dmart_deductions.reset_index(drop=True, inplace=True)
    dmart_deductions = None if dmart_deductions.empty else dmart_deductions
    
    final_dc[pdf_filename] = final_df
    return final_dc, dmart_deductions
