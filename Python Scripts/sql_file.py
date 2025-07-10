import pymssql
import pandas as pd
from math import ceil

def sql_formatter(dic):

    def chunk_list(lst, size):
        for i in range(0, len(lst), size):
            yield lst[i:i + size]

    conn_ = pymssql.connect(  # Remote Server Credentials
        host='',
        user=r'',
        password='',
        database=''
    )

    missing_invoices_list = []  # List to store missing invoices
    result_dict = {}
    max_batch_size = 200  # safe limit for SQL IN clause

    for key, value in dic.items():
        final = value
        invoice_ids = list(final['Invoice No.'])

        customers_list = []
        found_invoice_ids = set()

        for batch in chunk_list(invoice_ids, max_batch_size):
            # Proper escaping for OPENQUERY
            id_string = ", ".join(f"''{inv}''" for inv in batch)  # Note double single-quotes

            # Build inner query in one line, escaped
            inner_sql = f" "

            query = f"""
            
            """

            try:
                batch_customers = pd.read_sql_query(query, conn_)
                customers_list.append(batch_customers)
                found_invoice_ids.update(batch_customers['INVOICEID'].tolist())
            except Exception as e:
                print(f"Error in batch query for key {key}: {e}")
                continue

        # Combine all fetched data
        customers = pd.concat(customers_list, ignore_index=True) if customers_list else pd.DataFrame()

        # Identify missing invoice IDs
        all_invoice_ids = set(invoice_ids)
        missing_invoice_ids = all_invoice_ids - found_invoice_ids
        for invoice_id in missing_invoice_ids:
            missing_invoices_list.append({'Invoice No.': invoice_id})

        # Merge with original data
        final_data = pd.merge(final, customers, left_on="Invoice No.", right_on='INVOICEID', how="inner")
        final_data.drop_duplicates(subset=["Invoice No."], inplace=True)

        try:
            final_data.drop(["Page"], axis=1, inplace=True)
        except KeyError:
            pass

        df_grouped = final_data.groupby(["CUSTNO", "Payment Advice No."]).agg({
            "Invoice Amount": lambda x: [round(float(i), 2) for i in x],  
            "Payment Amount": lambda x: [round(float(i), 2) for i in x],  
            "TDS": lambda x: [round(float(i), 2) for i in x],  
            "Amount of Deduction": lambda x: [round(float(i), 2) for i in x],  
            "Invoice No.": lambda x: list(x),  
            "CUSTNAME": "first",  
            "INVOICEDATE": "min"  
        }).reset_index()

        df_grouped["Total"] = df_grouped["Payment Amount"].apply(lambda x: round(sum(x), 2))
        df_grouped["Total_TDS"] = df_grouped["TDS"].apply(lambda x: round(sum(x), 2))

        result_dict[key] = df_grouped

    conn_.close()

    # Create missing invoice DataFrame
    missing_invoices_df = pd.DataFrame(missing_invoices_list)
    missing_invoices_df = None if missing_invoices_df.empty else missing_invoices_df

    return result_dict, missing_invoices_df

