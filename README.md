![Python](https://img.shields.io/badge/python-3.10%2B-blue)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)
![Selenium](https://img.shields.io/badge/Selenium-Automation-green)
![pymssql](https://img.shields.io/badge/Database-SQLServer-blueviolet)
![Pandas](https://img.shields.io/badge/pandas-Data%20Processing-black?logo=pandas)

# Dynamics 365 Automation (Automated Invoice Processing & Journal Entry System)

This Python-based automation solution streamlines the processing of vendor invoices received via email attachments by automatically creating journal entries in Dynamics 365. It eliminates manual data entry and significantly accelerates the overall invoice processing workflow.

## Process Flow

1. Email Monitoring & File Organization
Continuously monitors incoming emails with the subject pattern "dynamics365 [vendor]" (e.g., dynamics365 dmart). Automatically downloads PDF attachments into dedicated vendor folders such as Reliance, Dmart, Lulu, and Aadhar.

2. Data Extraction & Validation
Automatically logs into the Dynamics 365 web portal.
Extracts key invoice details from PDFs, including amounts, descriptions, and vendor information.
Validates the extracted data by cross-referencing it with a remote database.
Categorizes records as matched or unmatched to guide subsequent processing steps.

3. Reporting & Documentation
Maintains comprehensive tracking of successful and failed transactions. Generates detailed Excel reports documenting successful entries along with a full audit trail.

4. Notification Handling
Sends summary emails upon completion with all relevant attachments—including PDFs, Excel reports, and data summaries. In case of errors or interruptions, failure notifications with detailed error information are promptly dispatched.

## Business Impact
- Cuts manual processing time by 80–90%
- Eliminates human errors in data entry
- Provides a thorough audit trail and transparent error reporting
- Enables same-day invoice processing and journal entry creation in Dynamics 365

## Tools & Technologies Used
- **openpyxl and xlsxwriter** — for creating and manipulating Excel files
- **selenium with chromedriver** — to automate web interactions with the Dynamics 365 portal
- **pandas** — for data cleaning, transformation, and formatting
- **pdfplumber, fitz, and PyMuPDF** — for extracting data from PDF invoices
- **win32com (win32)** — to automate email handling and interactions with Outlook
- **pymssql** — to connect and query the remote SQL database for data validation

## Prerequisites

- [Python 3.10 or higher](https://www.python.org/downloads/)
- Microsoft Outlook (installed and logged in)
- [ChromeDriver (compatible with your Chrome version)](https://developer.chrome.com/docs/chromedriver/downloads/)
- Virtual environment setup (recommended: `python -m venv VE`)

## Folder Structure

```
|-- Extracted_CSV     # Store Extacted Excels & Report Excels
|   |-- Aadhar    
|   |-- CSD
|   |-- Dmart
|   |-- LULU
|   |-- REPORT_Aadhar
|   |-- REPORT_Dmart
|   |-- REPORT_LULU
|   |-- REPORT_RIL
|   `-- RIL
|-- Failure_Files     # Store Failure PDF Files
|   |-- Aadhar
|   |-- CSD
|   |-- Dmart
|   |-- LULU
|   |-- MIL
|   |-- Other
|   `-- RIL
|-- Formats           # Store Email Attachments
|   |-- Aadhar
|   |-- CSD
|   |-- Dmart
|   |-- LULU
|   |-- MIL
|   |-- Other
|   `-- RIL
|-- Mailed_CSV        # Store Excels that are Mailed
|   |-- Aadhar
|   |-- CSD
|   |-- Dmart
|   |-- LULU
|   `-- RIL
|-- Processed         # Store PDFs that are Mailed
|   |-- Aadhar
|   |-- CSD
|   |-- Dmart
|   |-- LULU
|   `-- RIL
|-- Python Scripts    
|   |-- aadhar_extract.py     # Extraction 
|   |-- chromedriver.exe      # Chromedriver
|   |-- config.ini            # Config File
|   |-- dmart_extract.py      # Extraction
|   |-- download_file.py      # Download Mail Attachmets
|   |-- dynamics_login.py     # Login to Dynamics365
|   |-- lulu_extract.py       # Extraction
|   |-- mail_file.py          # Mail After Completion/Failure
|   |-- main.py               
|   |-- payment_auto.py       # Payment Automation
|   |-- ril_extract.py        # Extraction
|   |-- sql_file.py           # Remote DB Connection & Query
|   `-- tds_auto.py           # TDS Automation
|-- README.md
`-- requirements.txt
```

## Installation

1. **Clone the repository**

```bash
git clone https://github.com/chiragx16/Dynamics365_Automation.git
```

2. **Create and activate a virtual environment**

```bash
python -m venv VE
VE\Scripts\activate
```

3. **Install dependencies**

```bash
pip install -r requirements.txt
```

4. **Configure environment variables**

Create a `config.ini` file in the project Python Scripts folder with the following variables:

```
[Prod]
dynamics_email = 
dynamics_password = 
login_web =                    # Dynamics365 login URL
recipients = []    
main_path =                    # 'Formats' folder path
output_path =                  # 'Extracted_CSV' folder path
journal_web =                  # Dynamics365 journal URL
```

## Usage

After setup, run the following command to start automation:

```bash
python main.py
```
