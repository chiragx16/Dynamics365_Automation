from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.remote.webelement import WebElement
import time
import pandas as pd


import configparser

# Load the configuration
config = configparser.ConfigParser()
config.read('__path__') # config file path

# Read Config File section
journal_web = config.get('Prod', 'journal_web')




def dynamics_automation_tds(driver, df2, inv_type, control_acc):
    mail_df = pd.DataFrame(columns=df2.columns)
    colmn = ['invoice', 'mismatch']
    mismatch_df = pd.DataFrame(columns=colmn)
    not_available_list_tds = []  # DataFrame for missing invoices
    not_available_df_tds = pd.DataFrame(columns=['Missing Invoices in TDS Auto', "Payment Advice No."])

    # Get total number of rows
    total_rows = len(df2)
    
    # Group by "Payment Advice No."
    grouped = df2.groupby("Payment Advice No.")
    
    for payment_advice, group in grouped:
        
        journal_name = ""
        
        print(f"Doc No.: {payment_advice}")  # Print Payment Advice No. only once

        driver.get(journal_web)

        
        from selenium.webdriver.common.action_chains import ActionChains
        from selenium.webdriver.common.keys import Keys

        new_table_button = WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.ID, "ledgerjournaltable3_1_SystemDefinedNewButton"))
        )
        new_table_button.click()
        
        # Type Table name and open it
        file_open = WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, "//input[@role='combobox' and @aria-required='true' and @aria-label='Name']"))
        )
        file_open.send_keys("GLJV-U1")
        time.sleep(2)
    
        # Shift + Tab
        ActionChains(driver).key_down(Keys.SHIFT).send_keys(Keys.TAB).key_up(Keys.SHIFT).perform()
        time.sleep(1)
        ActionChains(driver).send_keys(Keys.ENTER).perform()
        time.sleep(4)
        
     
        
        journal = WebDriverWait(driver, 60).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'span.formCaption-context.largeViewSelector'))
        )
        journal_name = journal[1].text   
        print("Journal Name : " , journal_name)

        flag = 1
        credit_index = 0
        debit_index = 0
        
        for record, (index, row) in enumerate(group.iterrows(), start=1):
            present_tds = []
            try:
                print("-" * 70)
                print(f"  Index: {index}")
                print(f"  INVOICEACCOUNT: {row['CUSTNO']}")
                print(f"  INVOICEIDS: {row['Invoice No.']}")
                print(f"  TDS: {row['TDS']}")
                print(f"  TOTAL TDS: {row['Total_TDS']}")

                # time.sleep(5)
                # try:
                #     if flag == 1:
                #         # if first_time:
                #         date_cell = WebDriverWait(driver, 60).until(
                #             EC.presence_of_all_elements_located((
                #                 By.XPATH, "//input[contains(@id, 'LedgerJournalTrans_TransDate_') and contains(@id, '_input')]"
                #             ))
                #         )
                # except Exception as e:
                #     print("Date Exception")
                    
                # if date_cell:
                #     date_cell[-1].click()
                # else:
                #     print("No matching TransDate element found.")

                    
                # try:
                #     time.sleep(5)
                #     print("Trying to click Account type")
                #     # Enter Account type
                #     N = 3
                #     for _ in range(N):
                #         ActionChains(driver).send_keys(Keys.TAB).perform()
                #         time.sleep(0.5)
                #     ActionChains(driver).send_keys(Keys.DELETE).perform()
                #     time.sleep(2)
                #     ActionChains(driver).send_keys("Customer").perform()
                #     time.sleep(2)
                #     ActionChains(driver).send_keys(Keys.RETURN).perform()
                #     time.sleep(2)
                # except Exception as e:
                #     print("First error : ", e)

                partial_text = "Ledger"
                locator = (By.XPATH, f'//*[contains(@value, "{partial_text}") or contains(@title, "{partial_text}")]')
            
                # Use WebDriverWait to wait for at least one element matching the locator to be present
                WebDriverWait(driver, 60).until(
                    EC.presence_of_all_elements_located(locator)
                )
            
                elements: list[WebElement] = driver.find_elements(*locator)
            
                if elements:
                    first_element = elements[0]
                    driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", first_element)
                    time.sleep(1)
                    print(f"Found {len(elements)}")
                    first_element.send_keys(Keys.CONTROL + "a") 
                    time.sleep(1)
                    first_element.send_keys(Keys.BACKSPACE)
                    time.sleep(2)
                    first_element.click()
                    first_element.send_keys("Customer")
                    time.sleep(2)
                    first_element.send_keys(Keys.RETURN)
                    time.sleep(2)
                    print(f"Found {len(elements)} elements with value or title containing '{partial_text}'. Clicked the first one.")
                else:
                    print(f"No elements found with value or title containing '{partial_text}'.")


                print("Trying to enter Account value")
                # Enter Account value
                for _ in range(1):
                    ActionChains(driver).send_keys(Keys.TAB).perform()
                    time.sleep(0.5)
                invoice_acc = row['CUSTNO']
                ActionChains(driver).send_keys(invoice_acc).perform()
                time.sleep(5)
                ActionChains(driver).send_keys(Keys.RETURN).perform()
                time.sleep(2)
                ActionChains(driver).send_keys(Keys.RETURN).perform()
                time.sleep(3)

        
                print("Trying to enter Description")
                N = 1  # number of times you want to press TAB\n",

                for _ in range(N):
                    ActionChains(driver).send_keys(Keys.TAB).perform()
                    time.sleep(0.5)
                doc_no = row['Payment Advice No.']
                description_text = f"{inv_type} TDS Receivable - {doc_no}"
                print(f"Entering Description: {description_text}")
                ActionChains(driver).send_keys(description_text).perform()
                time.sleep(1)
        
        
                print("Trying to click Offeset Account")
                partial_text = "Ledger"
                locator = (By.XPATH, f'//*[contains(@value, "{partial_text}") or contains(@title, "{partial_text}")]')
            
                # Use WebDriverWait to wait for at least one element matching the locator to be present
                WebDriverWait(driver, 60).until(
                    EC.presence_of_all_elements_located(locator)
                )
            
                elements: list[WebElement] = driver.find_elements(*locator)
            
                if elements:
                    first_element = elements[0]
                    driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", first_element)
                    time.sleep(1)
                    print(f"Found {len(elements)}")
                    first_element.send_keys(Keys.CONTROL + "a") 
                    time.sleep(1)
                    first_element.send_keys(Keys.BACKSPACE)
                    time.sleep(2)
                    first_element.click()
                    first_element.send_keys("Customer")
                    time.sleep(2)
                    first_element.send_keys(Keys.RETURN)
                    time.sleep(2)
                    print(f"Found {len(elements)} elements with value or title containing '{partial_text}'. Clicked the first one.")
                else:
                    print(f"No elements found with value or title containing '{partial_text}'.")


        
                print("Trying to enter Offeset Account Type")
                # click offset account
                N = 1  # number of times you want to press TAB\n",
        
                for _ in range(N):
                    ActionChains(driver).send_keys(Keys.TAB).perform()
                    time.sleep(0.5)
                print(f"Entering control account: {control_acc}")
                ActionChains(driver).send_keys(control_acc).perform()
                time.sleep(5)
                ActionChains(driver).send_keys(Keys.RETURN).perform()
                time.sleep(2)
                ActionChains(driver).send_keys(Keys.RETURN).perform()
                time.sleep(3)
        
        
        
                N = 2  # number of times you want to press TAB\n",
                for _ in range(N):
                    ActionChains(driver).send_keys(Keys.TAB).perform()
                    time.sleep(0.5)
                    
        
                print("Trying to click Functions")
                # Click on "Functions" 
                find_func = WebDriverWait(driver, 60).until(EC.presence_of_all_elements_located((By.XPATH, "//*[contains(@id, 'LedgerJournalTransDaily') and contains(@id, '_LineFunctions_label')]")))
                time.sleep(5)
                find_func[-1].click()
        
                print("Trying to click Settle Transactions")

                # Click "Settle Transactions"
                time.sleep(4)
                settle_tran = WebDriverWait(driver, 60).until(
                    EC.presence_of_all_elements_located((By.XPATH, "//span[text()='Settle transactions']"))
                )
                time.sleep(3)
                settle_tran[-1].click()
        
                print("Trying to click Invoice Column")
                # Click on Invoice Column
                inv_filter = WebDriverWait(driver, 120).until(
                    EC.presence_of_all_elements_located((By.XPATH, "//*[contains(@id, 'Overview_CustTrans_Invoice') and contains(@id, '_header')]"))
                )
                inv_filter[0].click()

                try:
                    print("Trying to click filter dropdown")
                    # click filter dropdown
                    filter_dropdown = WebDriverWait(driver, 60).until(
                        EC.presence_of_all_elements_located((
                            By.XPATH, "//span[contains(@class, 'button-label-dropDown') and (text()='contains' or text()='is exactly' or text()='is not' or text()='does not contain' or text()='is one of' or text()='after' or text()='before' or text()='matches')]"
                        ))
                    )
                    filter_dropdown[-1].click()
        
                    print("Trying to click is one of option")
                    # select "is one of" option
                    filter_type = WebDriverWait(driver, 60).until(
                        EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'is one of')]"))
                    )
                    filter_type.click()
                except:
                    print(f"Dropdown filter click failed, skipping to Invoice Ids input.")

        
                print("Trying to Enter Invoice Ids")
                # Enter Invoice Ids
                time.sleep(5)
                inv_input = WebDriverWait(driver, 60).until(
                    EC.presence_of_all_elements_located((By.XPATH, "//*[contains(@name, 'FilterField_Overview_CustTrans_Invoice_Invoice_Input')]"))
                )
        
                inv_input[0].click()
                inv_li = eval(row['Invoice No.']) if isinstance(row['Invoice No.'], str) else row['Invoice No.']
                for i in inv_li:
                    inv_input[0].send_keys(str(i) + ',')
                inv_input[0].send_keys(Keys.RETURN)
                time.sleep(1)
    
                # Get the count of available rows dynamically using XPath
                xpath = "//*[contains(@id, 'MarkTrans') and contains(@id, '_container')]"
                available_rows = len(WebDriverWait(driver, 5).until(
                    EC.presence_of_all_elements_located((By.XPATH, xpath))
                ))
                
                print(f"Available Rows: {available_rows}")
        
                # Enter Prices
                price_li = eval(row['TDS']) if isinstance(row['TDS'], str) else row['TDS']
                for i in range(min(len(price_li), available_rows)):
                    # Clicks checkbox cell
                    xpath = f"//*[contains(@id, 'MarkTrans') and contains(@id, '{i}_container')]"
        
                    try:
                        # Wait for element to be present
                        cell_box = WebDriverWait(driver, 60).until(
                            EC.presence_of_all_elements_located((By.XPATH, xpath))
                        )
                        cell_box[0].click()
                    except:
                        continue
        
                    # Clicks actual checkbox
                    checkbox = WebDriverWait(driver, 60).until(
                        EC.presence_of_all_elements_located((By.CLASS_NAME, "_4ppbwx"))
                    )
                    checkbox[0].click()
        
                    from selenium.webdriver.common.action_chains import ActionChains
        
                    xpath = f"//*[contains(@id, 'Overview_CustTrans_Invoice_') and contains(@id, '{i}_input')]"
        
                    # Wait for element to be present
                    element1 = WebDriverWait(driver, 60).until(
                        EC.presence_of_all_elements_located((By.XPATH, xpath))
                    )
                    time.sleep(5)
        
                    if element1:
                        inv_id = element1[0]
                        value = inv_id.get_attribute('value')  

                        if value in inv_li:
                            index = inv_li.index(value)
                            present_tds.append(value)
                        else:
                            print("Not Found")
                            continue
        
                        print("Index : " , index, "Value : ", value)
                        print(present_tds)
        
                        from selenium.webdriver.common.keys import Keys
                        from selenium.webdriver.common.action_chains import ActionChains
        
                        N = 8  # number of times you want to press TAB\n",
        
                        actions = ActionChains(driver)
                        for _ in range(N):
                            actions = actions.send_keys(Keys.TAB)
                            actions.perform()
        
        
                        #clicks amount fields
                        from selenium.webdriver.common.action_chains import ActionChains
        
                        xpath = f"//*[contains(@id, 'settleField') and contains(@id, '{i}_input')]"
        

                        # List to store mismatch messages
                        mismatch_messages = []

                        element1 = WebDriverWait(driver, 60).until(
                            EC.presence_of_all_elements_located((By.XPATH, xpath))
                        )

                        action = ActionChains(driver)
                        action.double_click(element1[0]).perform()
                        time.sleep(5)
                        element1[0].send_keys(Keys.CONTROL, 'a')
                        element1[0].send_keys(Keys.BACKSPACE)
                        time.sleep(2)

                        entered_price = price_li[index]
                        element1[0].send_keys(entered_price)
                        time.sleep(2)

                        element1[0].send_keys(Keys.RETURN)
                        time.sleep(1)

                        # Get actual value shown in the field after entry
                        actual_value = element1[0].get_attribute("value").replace(",", "").strip()

                        # If mismatch, log it with invoice number
                        if float(actual_value.replace(",", "")) != entered_price:
                            invoice_no = inv_li[index]  # Use your existing invoice number list
                            message = (
                                f"Your amount is available {actual_value} but you are entering {entered_price} of full amount, so it's not a settled amount."
                            )


                            # Append to DataFrame
                            mismatch_df = pd.concat([
                                mismatch_df, 
                                pd.DataFrame([{'invoice': invoice_no, 'mismatch': message}])
                            ], ignore_index=True)
                        

                        from selenium.webdriver.common.keys import Keys
                        from selenium.webdriver.common.action_chains import ActionChains
        
                        N = 6  # number of times you want to press Shift + Tab
        
                        actions.key_down(Keys.SHIFT)  # Hold Shift
                        for _ in range(N):
                            actions.send_keys(Keys.TAB)  # Press Tab
                        actions.key_up(Keys.SHIFT)
                        actions.perform()
        
                        time.sleep(5)
        
                # Click OK Button
                ok_button = driver.find_element(By.NAME, "Save")
                ok_button.click()
                mail_df = pd.concat([mail_df, pd.DataFrame([row])], ignore_index=True)
                time.sleep(2)
                # Update the 'Journal' column for the corresponding 'Payment Advice No.'
                mail_df.loc[mail_df['Payment Advice No.'] == payment_advice, 'TDS_Journal'] = journal_name

                from selenium.webdriver.common.keys import Keys
                from selenium.webdriver.common.action_chains import ActionChains

                N = 8  # number of times you want to press Shift + Tab
                for _ in range(N):
                    actions.send_keys(Keys.TAB)  # Press Tab
                actions.perform()


                print("trying to get credit value" , credit_index)
                credit_xpath = f"//*[contains(@id, 'LedgerJournalTrans_AmountCurCredit_') and contains(@id, '{credit_index}_input')]"
                print("Found credit at ", credit_xpath)
                credit_element = WebDriverWait(driver, 60).until(
                    EC.presence_of_element_located((By.XPATH, credit_xpath))
                )
                credit_value = credit_element.get_attribute("value")

                print("Value : ", credit_value)

                credit_index += 2

                N = 5  # number of times you want to press Shift + Tab
        
                actions.key_down(Keys.SHIFT)  # Hold Shift
                for _ in range(N):
                    actions.send_keys(Keys.TAB)  # Press Tab
                actions.key_up(Keys.SHIFT)
                actions.perform()

                time.sleep(10)
                
                print("Flag before : ", flag)
                if flag == 1:
                    # Click on 'Account type' if visible and clickable
                    print("Clicking Accout type")
                    try:
                        account_type_element = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, "//div[normalize-space(text())='Account type']"))
                        )
                        driver.execute_script("arguments[0].scrollIntoView(true);", account_type_element)
                        time.sleep(1)
                        account_type_element.click()
                    except Exception as e:
                        print("cound not click account type")


                    try:
                        print("Clicking sort option")
                        # Click on 'Sort first to last' safely
                        sort_element = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, "//span[normalize-space(text())='Sort first to last']"))
                        )
                        driver.execute_script("arguments[0].scrollIntoView(true);", sort_element)
                        time.sleep(1)
                        sort_element.click()
                    except Exception as e:
                        print("cound not click account type")
        
                flag += 1

                print("Flag after : ", flag)
                
                print("Trying to click new button")
                time.sleep(5)
                new = WebDriverWait(driver, 60).until(
                    EC.presence_of_all_elements_located((By.XPATH, "//span[text()='New']"))
                )
                time.sleep(3)
                new[-1].click()  # Click the last "New" button found
                print("clicked new button")


                print("Trying to click Account type 2nd time")
                # Enter Account type
                partial_text = "Ledger"
                locator = (By.XPATH, f'//*[contains(@value, "{partial_text}") or contains(@title, "{partial_text}")]')
            
                # Use WebDriverWait to wait for at least one element matching the locator to be present
                WebDriverWait(driver, 60).until(
                    EC.presence_of_all_elements_located(locator)
                )
            
                elements: list[WebElement] = driver.find_elements(*locator)
            
                if elements:
                    first_element = elements[0]
                    driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", first_element)
                    time.sleep(1)
                    print(f"Found {len(elements)}")
                    first_element.send_keys(Keys.CONTROL + "a") 
                    time.sleep(1)
                    first_element.send_keys(Keys.BACKSPACE)
                    time.sleep(2)
                    first_element.click()
                    first_element.send_keys("Customer")
                    time.sleep(2)
                    first_element.send_keys(Keys.RETURN)
                    time.sleep(2)
                    print(f"Found {len(elements)} elements with value or title containing '{partial_text}'. Clicked the first one.")
                else:
                    print(f"No elements found with value or title containing '{partial_text}'.")

   
                print("Trying to enter Account value 2nd time")
                # Enter Account value
                for _ in range(1):  # N = 1
                    ActionChains(driver).send_keys(Keys.TAB).perform()
                    time.sleep(0.5)
                    
                invoice_acc = row['CUSTNO']
                print(f"Entering Account Number: {invoice_acc}")
                ActionChains(driver).send_keys(invoice_acc).perform()
                time.sleep(5)
                
                ActionChains(driver).send_keys(Keys.RETURN).perform()
                time.sleep(2)
                ActionChains(driver).send_keys(Keys.RETURN).perform()
                time.sleep(3)
        
                print("Trying to enter Description 2nd time")
                for _ in range(1):
                    ActionChains(driver).send_keys(Keys.TAB).perform()
                    time.sleep(0.5)

                # Enter Description
                doc_no = row['Payment Advice No.']
                description_text = f"{inv_type} TDS Receivable - {doc_no}"
                ActionChains(driver).send_keys(description_text).perform()
                time.sleep(1)
        
        
                for _ in range(1):
                    ActionChains(driver).send_keys(Keys.TAB).perform()
                    time.sleep(0.5)
                
                # Try entering the debit value

                print("Trying to enter debit value...")
                ActionChains(driver).send_keys(credit_value).perform()
                print("Debit entered successfully.")
                debit_index += 2
                

                print("Trying to click Offeset Account")
                partial_text = "Ledger"
                locator = (By.XPATH, f'//*[contains(@value, "{partial_text}") or contains(@title, "{partial_text}")]')
            
                # Use WebDriverWait to wait for at least one element matching the locator to be present
                WebDriverWait(driver, 60).until(
                    EC.presence_of_all_elements_located(locator)
                )
            
                elements: list[WebElement] = driver.find_elements(*locator)
            
                if elements:
                    first_element = elements[0]
                    driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", first_element)
                    time.sleep(1)
                    print(f"Found {len(elements)}")
                    first_element.send_keys(Keys.CONTROL + "a") 
                    time.sleep(1)
                    first_element.send_keys(Keys.BACKSPACE)
                    time.sleep(2)
                    first_element.click()
                    first_element.send_keys("Customer")
                    time.sleep(2)
                    first_element.send_keys(Keys.RETURN)
                    time.sleep(2)
                    print(f"Found {len(elements)} elements with value or title containing '{partial_text}'. Clicked the first one.")
                else:
                    print(f"No elements found with value or title containing '{partial_text}'.")


        
                print("Trying to enter Offeset Account Type")
                # click offset account
                N = 1  # number of times you want to press TAB\n",
        
                for _ in range(N):
                    ActionChains(driver).send_keys(Keys.TAB).perform()
                    time.sleep(0.5)
                print(f"Entering control account: {control_acc}")
                ActionChains(driver).send_keys(control_acc).perform()
                time.sleep(5)
                ActionChains(driver).send_keys(Keys.RETURN).perform()
                time.sleep(2)
                ActionChains(driver).send_keys(Keys.RETURN).perform()
                time.sleep(3)

                print("Trying to click new button")
                time.sleep(5)
                new = WebDriverWait(driver, 60).until(
                    EC.presence_of_all_elements_located((By.XPATH, "//span[text()='New']"))
                )
                time.sleep(3)
                new[-1].click()  # Click the last "New" button found
                print("clicked new button")
    
                    
            except Exception as e:
                print('=='*50)
                print(e)
                pass


            not_available_list_tds = [item for item in inv_li if item not in present_tds]  # Compute missing values in each iteration
            not_available_df_tds = pd.concat([not_available_df_tds, pd.DataFrame({'Missing Invoices in TDS Auto': not_available_list_tds, 'Payment Advice No.': payment_advice})], ignore_index=True)
        
        

        # not_available_list = list(set(invoice_numbers) - set(present_invs))
        # not_available_df = not_available_df.append({'Missing Invoice No.': c}, ignore_index=True)
        
    not_available_df_tds = None if not_available_df_tds.empty else not_available_df_tds
    mismatch_df = None if mismatch_df.empty else mismatch_df
    if not mismatch_messages:
        mismatch_messages = []  # Ensure it's an empty list if no mismatches

    return driver, mail_df, not_available_df_tds, mismatch_df