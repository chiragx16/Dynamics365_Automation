import os
import re
import pandas as pd
import pdfplumber


def extract_lulu(pdf_path):
    final_dc = {}
    pdf_filename = os.path.basename(pdf_path)
    with pdfplumber.open(pdf_path) as pdf:
        text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])
    
    # Extract Voucher/Cheque No.
    voucher_match = re.search(r"Voucher/Cheque\s*No\.?\s*[:\s]+(\d+)", text, re.IGNORECASE)
    voucher_cheque_no = voucher_match.group(1) if voucher_match else "Not Found"
    
    # Splitting Additions and Deductions Correctly
    deductions_start_match = re.search(r"Deductions Table:", text, re.IGNORECASE)
    if deductions_start_match:
        split_index = deductions_start_match.start()
        additions_text = text[:split_index]
        deductions_text = text[split_index:]
    else:
        additions_text = text
        deductions_text = ""
    
    # Regex to match transaction entries properly
    entry_pattern = re.compile(
        r"(\d{3})\s+([\w\s/.-]+?)\s+(\d+)\s+(\d{2}.\d{2}.\d{4})\s+([\w/-]+)\s+(\d{2}.\d{2}.\d{4})\s+([\d,]+.\d{2})",
        re.DOTALL
    )
    
    def extract_data(text_section):
        lines = text_section.split("\n")
        extracted_data = []
        buffer = []
    
        for line in lines:
            if re.match(r"^\s*\d{3}\s+", line):
                if buffer:
                    extracted_data.append(" ".join(buffer).strip())
                buffer = [line]
            else:
                buffer.append(line)
    
        if buffer:
            extracted_data.append(" ".join(buffer).strip())
    
        clean_data = []
        for entry in extracted_data:
            match = entry_pattern.match(entry)
            if match:
                sl_no, tran_type, doc_no, inv_dt, inv_no, post_dt, amount = match.groups()
                clean_data.append((
                    voucher_cheque_no, 
                    sl_no.strip(),
                    tran_type.strip(),
                    doc_no.strip(),
                    inv_dt.strip(),
                    inv_no.strip(),
                    post_dt.strip(),
                    float(amount.replace(",", ""))
                ))
                
        list1 = []
        list2 = []
        current_list = list1
    
        for item in clean_data:
            if item[1] == '001' and current_list is list1 and list1:
                current_list = list2  # Switch to list2 only after the first segment completes
            current_list.append(item)
    
        return list1, list2
    
    additions, deductions = extract_data(additions_text)
    
    additions_columns = ["Payment Advice No.", "Sl No.", "Tran. Type", "Doc No.", "Inv/Ref Dt", "Invoice No.", "Posting Dt", "Payment Amount"]
    deductions_columns = ["Payment Advice No.", "Sl No.", "Tran. Type", "Doc No.", "Inv/Ref Dt", "Invoice No.", "Posting Dt", "Payment Amount"]
    
    additions_df = pd.DataFrame(additions, columns=additions_columns)
    deductions_df = pd.DataFrame(deductions, columns=deductions_columns)
    
    def get_deduction_amount(invoice_no):
        return deductions_df.loc[deductions_df['Invoice No.'] == invoice_no, 'Payment Amount'].sum() if invoice_no in deductions_df['Invoice No.'].values else 0
    
    additions_df["Invoice Amount"] = additions_df["Payment Amount"]  
    additions_df["TDS"] = 0.0 
    additions_df["Amount of Deduction"] = additions_df["Invoice Amount"] - additions_df["Payment Amount"]
    additions_df['Other Deductions'] = additions_df['Invoice No.'].apply(get_deduction_amount)
    # additions_df["Other Amount"] = 0.0

    
    additions_total = additions_df["Payment Amount"].sum()
    deductions_total = deductions_df["Payment Amount"].sum()
    net_total = additions_total - deductions_total

    if deductions_df is not None or not deductions_df.empty:
            deductions_df.reset_index(drop=True, inplace=True)

    deductions_df = None if deductions_df.empty else deductions_df


    if additions_df is not None or not additions_df.empty:
            additions_df.reset_index(drop=True, inplace=True)


    final_dc[pdf_filename] = additions_df
    # print(additions_df.to_string())
    # print(deductions_df.to_string())
    return final_dc, deductions_df
