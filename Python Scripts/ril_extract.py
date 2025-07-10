import os
import re
import fitz  # PyMuPDF (for PDF processing)
import pandas as pd

def extract_invoice_descriptions(text):
    """Extracts invoice descriptions including standard, date-based, new TI-XXXXX patterns, and TDS separately."""
    # print(repr(text))

    # Invoice pattern (ex: SIU2A2122-18109 and SIUIS2122-09526)
    invoice_pattern = re.compile(r"([A-Z]{3,4}\d[A-Z]?\d{4,6}-\d+|[A-Z]{5}\d{4}-\d{5})")  

    # New Date-based Pattern (Ignoring 8-digit number)
    date_based_pattern = re.compile(r"\n\d{2}/\d+\n")  

    # New Pattern (e.g., 1030257748\n1)
    new_number_pattern = re.compile(r"(\d{5,})\n(\d+)\n")  

    # TI-XXXXX Invoice Pattern (ex: TI-0001807)
    ti_invoice_pattern = re.compile(r"([A-Z]{2}-\d{6,})")  
    # Alternate pattern for TI invoices with slashes and optional letters (e.g., TI-A000180/23)
    ti_invoice_pattern2 = re.compile(r"[A-Z]{2}-[A-Z]?\d{6,}/\d+")  

    # Stop Phrases to remove unwanted text from description (extra text from begining and end of page)
    stop_phrases = [
        "Your A/c with us", 
        "Bikaji Foods International Ltd"
    ]  
    number_pattern = r"^\d+(\.\d+)?$"  # Matches standalone numbers (amounts)
    date_pattern = r"^\d{2}\.\d{2}\.\d{4}$"  # Matches standalone dates (DD.MM.YYYY)

    # Find all occurrences
    standard_matches = list(invoice_pattern.finditer(text))
    date_based_matches = list(date_based_pattern.finditer(text))
    new_number_matches = list(new_number_pattern.finditer(text))
    ti_matches = list(ti_invoice_pattern.finditer(text))  
    ti_matches2 = list(ti_invoice_pattern2.finditer(text))

    # Combine all matches and sort them based on their position in the document
    all_matches = standard_matches + date_based_matches + new_number_matches + ti_matches + ti_matches2
    all_matches.sort(key=lambda match: match.start())  # Ensure correct order in text

    invoice_entries = []  # Store extracted data

    # Iterate through each matched invoice ID and extract the associated description
    for i in range(len(all_matches)):
        match = all_matches[i]
        
        # Extract appropriate invoice number format
        if match in new_number_matches: # Use the second group for new number-based invoices; else use full match
            invoice_id = match.group(2) 
        else:
            invoice_id = match.group().strip()

        start_pos = match.end()  # Start after the match
        end_pos = all_matches[i + 1].start() if i + 1 < len(all_matches) else len(text)  

        # Extract description text between two invoice entries 
        description_text = text[start_pos:end_pos].strip()

        # **Processing the Description**
        description_lines = description_text.split("\n")
        filtered_lines = []
        prev_was_number = False   # Flag to track consecutive amounts

        for idx, line in enumerate(description_lines):
            stripped_line = line.strip()

            # Skip lines with date-only values
            if re.match(date_pattern, stripped_line):
                continue  
            
            # Remove standalone amounts (unless followed by a date)
            if re.match(number_pattern, stripped_line):
                if prev_was_number and filtered_lines:
                    filtered_lines.pop()  # Remove previous amount if both are amounts
                    prev_was_number = False  
                    continue  

                # Retain the amount if followed by a date
                if (idx + 1 < len(description_lines)) and re.match(date_pattern, description_lines[idx + 1].strip()):
                    if not prev_was_number:  
                        filtered_lines.append(stripped_line)  

                    prev_was_number = False  
                    continue  

                else:
                    prev_was_number = True  
                    continue  

            filtered_lines.append(stripped_line)  # Add valid description line

        # **Fix: Strip extra spaces and new lines from description**
        description_text = "\n".join(filtered_lines).strip()

        # Identify and extract any TDS-related lines separately from description
        tds_pattern = re.compile(r"(\(TDS[^\n]*)", re.IGNORECASE)  
        tds_match = tds_pattern.search(description_text)

        tds_text = None
        if tds_match:
            tds_text = tds_match.group(1).strip()
            description_text = description_text.split(tds_text)[0].strip()  # Remove TDS part from main description

        # Stop parsing description if certain phrases are encountered (e.g., page footers, unwanted noise)
        stop_phrases_regex = "|".join(map(re.escape, stop_phrases + ["\nTotal", "\nBIKAJI FOODS"]))
        description_text = re.split(stop_phrases_regex, description_text, maxsplit=1)[0].strip()

        # Fix spacing issues in numeric descriptions (e.g., "54.01-)" should be spaced properly like "54.01-    )" 
        description_text = re.sub(r"(\d+\.\d+-)(?!\s{3})", r"\1   ", description_text)
        if tds_text:
            tds_text = re.sub(r"(\d+\.\d+-)(?!\s{3})", r"\1   ", tds_text)

        # Append the cleaned description for the invoice
        invoice_entries.append({"Invoice": invoice_id, "Description": description_text})

        # If TDS content is found, add it as a separate entry linked to the same invoice ID
        if tds_text:
            invoice_entries.append({"Invoice": invoice_id, "Description": tds_text})

#     print("\n" , invoice_entries)
    return invoice_entries




def add_hyphen(code):
    """Ensure the invoice number has a hyphen in the correct place."""
    if '-' not in code[-6:]:
        return code[:-5] + '-' + code[-5:]
    return code

def add_hyphen_conditionally(invoice_no):
    """Apply add_hyphen only if the invoice number does not match the pattern."""
    pattern = r"[A-Z]{2}-\d{6,}"
    
    if re.match(pattern, str(invoice_no)):  
        return invoice_no  # Return as is if it matches the pattern (if it already has "-" )
    else:
        return add_hyphen(invoice_no)  # Apply add_hyphen otherwise




def extract_amounts(text):
    """Extract invoice IDs, amounts, and TDS from the text."""
    
    id_patterns = [
        r"\n[A-Z]{3,4}\d[A-Z]?\d{4,6}-\d+\n",  # ex: "\nSIU2A2122-18109\n"
        r"\n[A-Z]{2}-\d{6,}\n",  # ex: "TI-123456"
        r"\n[A-Z]{2}-[A-Z]?\d{6,}/\d+\n", # ex: "\nTI-A000180/23\n"
        r"\n[A-Z]{5}\d{4}-\d{5}\n" # ex: "SIUIS2122-09526"

    ]
    
    amt_pattern = r"Rs\.(\d+\.\d+)Amtwithtax"
    tds_pattern = r"TDSAmount(\d+\.\d+)-"

    # Find all invoice IDs using multiple patterns
    id_positions = []
    for pattern in id_patterns:
        id_positions.extend((match.start(), match.group()) for match in re.finditer(pattern, text))

    # Find amounts and TDS
    amt_matches = [(match.start(), match.group(1)) for match in re.finditer(amt_pattern, text)]
    tds_matches = [(match.start(), match.group(1)) for match in re.finditer(tds_pattern, text)]

    results = {}

    # Assign amounts to the closest preceding invoice ID
    for match_pos, amount in amt_matches:
        closest_id = None
        for pos, id_value in id_positions:
            if pos < match_pos:
                closest_id = id_value
            else:
                break  

        if closest_id:
            if closest_id not in results:
                results[closest_id] = {"amount": None, "tds": None}  
            results[closest_id]["amount"] = amount

    # Assign TDS to the closest preceding invoice ID
    for match_pos, tds in tds_matches:
        closest_id = None
        for pos, id_value in id_positions:
            if pos < match_pos:
                closest_id = id_value
            else:
                break  

        if closest_id:
            if closest_id not in results:
                results[closest_id] = {"amount": None, "tds": None}  
            results[closest_id]["tds"] = tds  

    return [(key.strip(), values["amount"], values["tds"]) for key, values in results.items()]








def extract_text_from_pdf_ril(pdf_path):
    """Process a single PDF file and return extracted data as a DataFrame."""
    final_data = []
    amount_tds = []
    descriptions = []  # Store descriptions separately

    final_dc = {}
    # Patterns to extract whole invoice lines (with Amounts)
    invoice_patterns = [
        r"[A-Z]{3}\d[A-Z]?\d{4,6}-\d{5}\n\d+\.\d+\n\d+\.\d+",  
        r"[A-Z]{3}\d+[A-Z]+\d+-\d+\.\n\d+\.\d+\n\d+\.\d+",  
        r"[A-Z]{3}\d[A-Z]\d+/\d\n\d+\.\d+\n\d+\.\d+",  
        r"[A-Z]{2}-\d{6,}\n\d+\.\d+\n\d+\.\d+",
        r"[A-Z]{2}-[A-Z]?\d{6,}/\d+\n\d+\.\d+\n\d+\.\d+",
        r"[A-Z]{5}\d{4}-\d{5}\n\d+\.\d+\n\d+\.\d+"
    ]

    doc_pattern = r"Doc\.number\s*(\d+)/\d+"  

    pdf_filename = os.path.basename(pdf_path)

    document = fitz.open(pdf_path)
    last_doc_number = None  

    extracted_data = []  

    last_invoice_id = None 

    for page_num in range(document.page_count):
        page = document.load_page(page_num)
        text = page.get_text("text")
        text = re.sub(r"(-\d+)\.", r"\1", text) # remove "." from text for Malformed Invoice like SIU1N2122-16493.
        text = re.sub(r"(-\d+)/", r"\1", text) # remove "." from text for Malformed Invoice like SIU2A2122-12709/
        text = text.replace(" ", "").replace(",", "").replace("_", "").replace("`","") # remove (`) from Invoice like `SIU1S2122-04518
        
        text2 = page.get_text("text").replace(",", "").replace("_", "")

        # if page_num == 5:
        #     print(repr(text))
        #     print("\n")

            # print("\n" , last_invoice_id , page_num)
            

        # Check if the page starts with TDS
        tds_at_start_pattern = re.search(r"Narration\.\s*\n*\(TDSAmount(\d+\.\d+)-\)", text)
        # if TDS found at the start of page assign it to the last invoice of previous page
        if tds_at_start_pattern and last_invoice_id:
            tds_value = tds_at_start_pattern.group(1)
            # print("\nTDS Value : " , tds_value , "\n")
            amount_tds.append((last_invoice_id, None, tds_value))



        
        # extract Description/Deductions
        invoice_data = extract_invoice_descriptions(text2)
        
        for entry in invoice_data:
            descriptions.append((entry["Invoice"], entry["Description"]))
            last_invoice_id = entry["Invoice"]  # Update last seen invoice


        ids = extract_amounts(text)
        if ids:
            amount_tds += ids
            # last_invoice_id = ids[-1][0] 

        # if page_num == 5:
        #     print("\nAmount TDS : " , amount_tds , "\n")
        

        # print("\n" , last_invoice_id)

        # print("\nTDS : " , amount_tds)

        doc_match = re.search(doc_pattern, text)
        if doc_match:
            last_doc_number = doc_match.group(1)  

        page_matches = []
        for pattern in invoice_patterns:
            page_matches.extend(re.findall(pattern, text))

        unique_matches = list(set(page_matches))
        page_data = [entry.split('\n') for entry in unique_matches]

        if page_data:
            df = pd.DataFrame(page_data, columns=["Invoice No.", "Invoice Amount", "Payment Amount"])
            df["Filename"] = pdf_path  
            df["Payment Advice No."] = last_doc_number  

            df["Invoice No."] = df["Invoice No."].str.replace('.', '', regex=False) # remove '.' from invoice
            df["Invoice No."] = df["Invoice No."].str.replace('/', '', regex=False) # remove '/' from invoice
            df["Invoice No."] = df["Invoice No."].apply(add_hyphen_conditionally)
            extracted_data.append(df)

    if extracted_data:
        final_df = pd.concat(extracted_data, ignore_index=True)
        final_df = final_df[["Invoice No.", "Invoice Amount", "Payment Amount", "Filename", "Payment Advice No."]]
        final_df = final_df.drop_duplicates(subset=["Invoice No.", "Payment Advice No."])

        final_df["Invoice Amount"] = pd.to_numeric(final_df["Invoice Amount"], errors='coerce')
        final_df["Payment Amount"] = pd.to_numeric(final_df["Payment Amount"], errors='coerce')

        final_df["Amount of Deduction"] = final_df["Invoice Amount"] - final_df["Payment Amount"]

        amount_tds_df = pd.DataFrame([(inv, tds) for inv, _, tds in amount_tds], columns=["Invoice No.", "TDS"])

        amount_tds_df["TDS"] = pd.to_numeric(amount_tds_df["TDS"], errors='coerce')

        final_df = final_df.merge(amount_tds_df, on="Invoice No.", how="left")
        final_df = final_df.drop_duplicates(subset=["Invoice No.", "Payment Advice No."])

        final_df[["TDS"]] = final_df[["TDS"]].fillna(0)
        
        descriptions_df = pd.DataFrame(descriptions, columns=["Invoice No.", "Description"])
        descriptions_df = descriptions_df[descriptions_df["Description"].str.strip() != ""]
        
        descriptions_df.reset_index(drop=True, inplace=True)
        descriptions_df = None if descriptions_df.empty else descriptions_df
        
        final_df.reset_index(drop=True, inplace=True)
        final_dc[pdf_filename] = final_df
        descriptions_df.to_excel("des.xlsx")
        final_df.to_excel("final.xlsx")
        
        return final_dc, descriptions_df
    else:
        print("No invoice data found in the PDF.")
        return None
