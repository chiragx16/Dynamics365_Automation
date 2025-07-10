import os
import re
import pandas as pd
import pdfplumber



def extract_aadhar_format(pdf_path):
        
    # Initialize lists
    final_dc = {}
    all_tables = []
    document_numbers = []
    document_dates = []
    document_no = []
    reference_numbers = []
    invoice_dates = []
    site = []
    dr_cr = []
    invoice_amounts = []
    tds_amount = []
    net_amount = []
    
    pdf_filename = os.path.basename(pdf_path)
    
    # Regular expression for matching date formats (DD.MM.YYYY, DD-MM-YYYY)
    date_pattern = r"\b\d{1,2}[./-]\d{1,2}[./-]\d{4}\b"

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):  # Loop through pages
            tables = page.extract_tables()  # Extract tables
            
            for table in tables:  # Loop through tables
                df = pd.DataFrame(table)  # Convert table to DataFrame
                all_tables.append(df)  # Store table
                
                # Iterate through each row and column
                for col in range(df.shape[1]):  
                    for row in range(df.shape[0]):  
                        cell_value = df.iloc[row, col]  # Get cell value
                        
                        # Search for "Document No : 12345" pattern
                        if cell_value and re.search(r'Document\s*No\s*:\s*(\d+)', cell_value, re.IGNORECASE):
                            match = re.search(r'Document\s*No\s*:\s*(\d+)', cell_value)
                            if match:
                                document_no.append(match.group(1))
                        
                        # Search for "Document\nNumber" and extract numbers below
                        if cell_value and "document" in cell_value.lower() and "number" in cell_value.lower():
                            for below_row in range(row + 1, df.shape[0]):
                                below_cell = df.iloc[below_row, col]
                                if below_cell and re.fullmatch(r"\d+", below_cell.strip()):  
                                    document_numbers.append(below_cell.strip())
    
                        # Search for "Document Date" and extract date below
                        if cell_value and "document\ndate" in cell_value.lower():
                            for below_row in range(row + 1, df.shape[0]):
                                below_cell = df.iloc[below_row, col]
                                if below_cell and re.search(date_pattern, below_cell):  
                                    date_match = re.search(date_pattern, below_cell)
                                    document_dates.append(date_match.group(0))  
    
                        # Search for "Reference No" and extract values below until None is found
                        if cell_value and "reference" in cell_value.lower() and "no" in cell_value.lower():
                            for below_row in range(row + 1, df.shape[0]):
                                below_cell = df.iloc[below_row, col]
                                if below_cell and below_cell.strip():  # Stop when None or empty cell is found
                                    reference_numbers.append(below_cell.replace("\n", " ").strip())
                                else:
                                    break  # Stop when an empty cell is found
    
                        # Search for "Invoice\nDate" and extract values below until None is found
                        if cell_value and "invoice\ndate" in cell_value.lower():
                            for below_row in range(row + 1, df.shape[0]):
                                below_cell = df.iloc[below_row, col]
                                if below_cell and re.search(date_pattern, below_cell):  
                                    date_match = re.search(date_pattern, below_cell)
                                    invoice_dates.append(date_match.group(0)) 
    
                        # Search for "Site" and extract values below until None is found
                        if cell_value and "site" in cell_value.lower():
                            for below_row in range(row + 1, df.shape[0]):
                                below_cell = df.iloc[below_row, col]
                                if below_cell and below_cell.strip():  # Stop when None or empty cell is found
                                    site.append(below_cell.strip())
                                else:
                                    break  # Stop when an empty cell is found
                                    
                        # Search for "Dr/Cr" and extract values below until None is found
                        if cell_value and "dr/cr" in cell_value.lower():
                            for below_row in range(row + 1, df.shape[0]):
                                below_cell = df.iloc[below_row, col]
                                if below_cell and below_cell.strip():  # Stop when None or empty cell is found
                                    dr_cr.append(below_cell.strip())
                                else:
                                    break  # Stop when an empty cell is found
    
                        # Search for "Invoice\nAmount" and extract values below until None is found
                        if cell_value and "invoice\namount" in cell_value.lower():
                            # print(f"Found 'Document Date' at Page {page_num}, Row {row}, Col {col}")
                            for below_row in range(row + 1, df.shape[0]):
                                below_cell = df.iloc[below_row, col]
                                if below_cell and below_cell.strip():  # Stop when None or empty cell is found
                                    invoice_amounts.append(below_cell.strip())
                                else:
                                    break  # Stop when an empty cell is found
    
                        # Search for "TDS Amount" and extract values below until None is found
                        if cell_value and "tds amount" in cell_value.lower():
                            # print(f"Found 'Document Date' at Page {page_num}, Row {row}, Col {col}")
                            for below_row in range(row + 1, df.shape[0]):
                                below_cell = df.iloc[below_row, col]
                                if below_cell and below_cell.strip():  # Stop when None or empty cell is found
                                    tds_amount.append(below_cell.strip())
                                else:
                                    break  # Stop when an empty cell is found
    
                        # Search for "Net Amount" and extract values below until None is found
                        if cell_value and "net amount" in cell_value.lower():
                            # print(f"Found 'Document Date' at Page {page_num}, Row {row}, Col {col}")
                            for below_row in range(row + 1, df.shape[0]):
                                below_cell = df.iloc[below_row, col]
                                if below_cell and below_cell.strip():  # Stop when None or empty cell is found
                                    net_amount.append(below_cell.strip())
                                else:
                                    break  # Stop when an empty cell is found

        # covert column to digits                 
        invoice_amounts = pd.to_numeric(invoice_amounts, errors='coerce')
        net_amount = pd.to_numeric(net_amount, errors='coerce')
            
        # Create DataFrame
        final_df = pd.DataFrame({
            "Payment Advice No.": [document_no[0] if document_no else ""] * len(document_numbers),
            "Document Numbers": document_numbers,
            "Document Dates": document_dates,
            "Invoice No.": reference_numbers,
            "Invoice Dates": invoice_dates,
            "Site": site,
            "Dr/Cr": dr_cr,
            "Invoice Amount": invoice_amounts,
            "TDS": tds_amount,
            "Payment Amount": net_amount,
            # "Other Amount": 0.0,
            "Amount of Deduction": invoice_amounts - net_amount
        })
        final_df["Invoice No."] = final_df["Invoice No."].str.replace(" ", "")

        pattern = r"^[A-Za-z]{3}\d+[A-Za-z]\d*-\d+$" # invoice pattern

        # Filter rows where "Invoice No." matches the pattern
        aadhar_payment_df = final_df[final_df["Invoice No."].str.match(pattern, na=False)]
        if not aadhar_payment_df.empty:
            aadhar_payment_df.reset_index(drop=True, inplace=True)

        
        # Create Deduction DF where Invoice No does not match pattern
        aadhar_deductions_df = final_df[~final_df["Invoice No."].str.match(pattern, na=False)]
        if aadhar_deductions_df is not None or not aadhar_deductions_df.empty:
            aadhar_deductions_df.reset_index(drop=True, inplace=True)

        # return 'None' if no deductions in PDF
        aadhar_deductions_df = None if aadhar_deductions_df.empty else aadhar_deductions_df
        
        # display(aadhar_deductions_df)

    # print(final_df.to_string())
    final_dc[pdf_filename] = aadhar_payment_df
    return final_dc, aadhar_deductions_df
