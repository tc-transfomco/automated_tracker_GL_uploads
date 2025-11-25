import os
import sys
from datetime import datetime, date
from typing import Any, List
from pathlib import Path
import pandas as pd
import traceback
from mail_util import MailUtil


THIRD_PARTY_REPLACEMENT_PURCHASES = 'TRP'     
SAMSUNG_TO_LIST = "samsung_to_list" 


ACCOUNT_NAME = 'ACCOUNT NAME'
ACCOUNT_NUMBER = 'ACCOUNT NUMBER'
EMPLOYEE_ID = 'EMPLOYEE ID'
TRANSACTION_DATE = 'TRANSACTION DATE'
POSTING_DATE = 'POSTING DATE'
MERCHANT_NAME = 'MERCHANT NAME'
MCC = 'MCC'
TRANSACTION_AMOUNT = 'TRANSACTION AMOUNT'
COST_CENTER = 'COST CENTER'
GL = 'GL'
DIVISION = 'DIVISION'
TRANSFORM_ONECARD_CB_NAM_US = 'TRANSFORM ONECARD CB NAM US'
COMPANY_NUMBER = 'Company Number'
REFERENCES_NUMBER = 'References number'
AUTHNUMBER = 'AuthNumber'
CUST_CHG_AMOUNT = 'Cust Chg Amount'
COST_CENTER = 'Cost Center'
TRANSACTION_AMOUNT = 'Transaction Amount'
AGREEMENT_TYPE = 'Agreement Type'
STATE = 'State'
DOC_NBR = 'Doc NBR'
OFFSET_AMOUNT = 'OffSet Amount'
COST_CENTER2 = 'Cost Center2'
COGS = 'COGS'
AMT_50905 = '50905 Amt'
JE_YN = 'JE (Y/N)'
COLUMN6 = 'Column6'
COLUMN7 = 'Column7'
COLUMN8 = 'Column8'

LEN_OPERATING_UNIT = 5
LEN_DIVISION = 4
LEN_ACCOUNT = 5
LEN_SOURCE = 3
LEN_CATEGORY = 4
LEN_YEAR = 4
LEN_WEEK = 2
LEN_COST = 17
LEN_CURRENCY_CODE = 3
LEN_SELLING_VALUE = 17
LEN_STATISTICAL_AMOUNT = 17
LEN_STATISTICAL_CODE = 3
LEN_REVERSAL_FLAG = 1
LEN_REVERSAL_YEAR = 4
LEN_REVERSAL_WEEK = 2
LEN_BACKTRAFFIC_FLAG = 1
LEN_SYSTEM_SWITCH = 2
LEN_DESCRIPTION = 30
LEN_ENTRY_TYPE = 3
LEN_RECORD_TYPE = 2
LEN_REF_NBR_1 = 10
LEN_DOC_NBR = 15
LEN_REF_NBR_2 = 10
LEN_MISC_1 = 20
LEN_MISC_2 = 20
LEN_MISC_3 = 20
LEN_TO_FROM = 6
LEN_DOC_DATE = 8
LEN_EXP_CODE = 3
LEN_EMP_NBR = 7
LEN_DET_TRAN_DATE = 6
LEN_ORIG_ENTRY = 15
LEN_REP_FLG = 1
LEN_ORU = 6
LEN_GL_TRXN_DATE = 8
LEN_FILLER = 16

SANDBOX_DESTINATION = "//SHSNGSTFSX/shs_boomi_vol/Test/GL_Files/300-byte"
# SANDBOX_DESTINATION = "\\SHSNGSTFSX\shs_boomi_vol\Test\GL_Files\300-byte"
PROD_PATH = '//SHSNGPRFSX/shs_boomi_vol/GL_Files/300-byte'
# PROD_PATH = '\\SHSNGPRFSX\shs_boomi_vol\GL_Files\300-byte'

# Pick up all files that don't have _completed in the name. Then throw _completed in the name. 
PICKUP_LOCATION = Path("//uskihfil5.kih.kmart.com/workgrp/Finance/AccountsPayable/Samsung Re-Class")
# \\uskihfil5.kih.kmart.com\workgrp\Finance\AccountsPayable\Samsung Re-Class

OPERATING_UNIT = 'operating unit'
# DIVISION =
ACCOUNT = 'account'
COST = 'cost'
# DOC_NBR =
# REFERENCES_NUMBER =
# MERCHANT_NAME =
# ACCOUNT_NAME =
DET_TRAN_DATE = 'det tran date'

class OpenFileException(Exception):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)


def justify(input, amount, fillchar=' '):
    input = str(input).replace("\n", " ").replace("\t", " ")
    return input[:amount].rjust(amount, fillchar)


def process_file(file_path: Path):
    def process_row(row):
        do_cogs = int(row[CUST_CHG_AMOUNT]) != 0             # if there is a custom charge amount, we have to do a cogs row. 
        
        main_entry = {}
        offset_entry = {}
        cogs_entry = {}


        main_entry[POSTING_DATE] = _get_gl_trxn_date(datetime.strptime(row[POSTING_DATE], "%m/%d/%Y"))
        offset_entry[POSTING_DATE] = _get_gl_trxn_date(datetime.strptime(row[POSTING_DATE], "%m/%d/%Y"))
        if do_cogs:
            cogs_entry[POSTING_DATE] = _get_gl_trxn_date(datetime.strptime(row[POSTING_DATE], "%m/%d/%Y"))

        # This next line is only for email formatting. 
        main_entry[CUST_CHG_AMOUNT] = row[CUST_CHG_AMOUNT]                
        
        main_entry[OPERATING_UNIT] = row[COST_CENTER] 
        offset_entry[OPERATING_UNIT] = 4542 
        if do_cogs:        
            cogs_entry[OPERATING_UNIT] = row[COST_CENTER] 

        # DIVISION is just hard coded, not taken from division in the file. 
        main_entry[DIVISION] = 400
        offset_entry[DIVISION] = 530
        if do_cogs:        
            cogs_entry[DIVISION] = 400

        main_entry[ACCOUNT] = 50905 
        offset_entry[ACCOUNT] = 63005
        if do_cogs:        
            cogs_entry[ACCOUNT] = 50195 
        
        main_entry[COST] = row[AMT_50905] 
        offset_entry[COST] = row[TRANSACTION_AMOUNT] * -1 
        if do_cogs:        
            cogs_entry[COST] = round(row[CUST_CHG_AMOUNT] * 0.75, 2) 

        # If the doc number has issues, don't process this record and sent 
        row_doc_num:str = str(int(float(row[DOC_NBR]))).strip()
        if len(row_doc_num) != 12 or not row_doc_num.isdigit():
            return None, None, None, True
        main_entry[DOC_NBR] = row_doc_num
        offset_entry[DOC_NBR] = row_doc_num
        if do_cogs:        
            cogs_entry[DOC_NBR] = row_doc_num

        main_entry[REFERENCES_NUMBER] = row[REFERENCES_NUMBER] 
        offset_entry[REFERENCES_NUMBER] = row[REFERENCES_NUMBER]
        if do_cogs:        
            cogs_entry[REFERENCES_NUMBER] = row[REFERENCES_NUMBER] 

        main_entry[MERCHANT_NAME] = row[MERCHANT_NAME] 
        offset_entry[MERCHANT_NAME] = row[MERCHANT_NAME]
        if do_cogs:        
            cogs_entry[MERCHANT_NAME] = row[MERCHANT_NAME] 

        main_entry[ACCOUNT_NUMBER] = row[ACCOUNT_NUMBER] 
        offset_entry[ACCOUNT_NUMBER] = row[ACCOUNT_NUMBER]
        if do_cogs:        
            cogs_entry[ACCOUNT_NUMBER] = row[ACCOUNT_NUMBER] 

        main_entry[DET_TRAN_DATE] = row[POSTING_DATE] 
        offset_entry[DET_TRAN_DATE] = row[POSTING_DATE]
        if do_cogs:        
            cogs_entry[DET_TRAN_DATE] = row[POSTING_DATE] 
        return main_entry, offset_entry, cogs_entry, False

    def fix_money_row(orig_df, row_name):
        orig_df[row_name] = orig_df[row_name].astype(str).str.replace("$", "", regex=False)
        orig_df[row_name] = orig_df[row_name].astype(str).str.replace(",", "",regex=False)
        orig_df[row_name] = pd.to_numeric(orig_df[row_name], errors='raise')
        return orig_df

    file_name = file_path.stem

    df = pd.read_excel(file_path)
    df.columns = df.columns.str.strip()

    money_rows = [AMT_50905, TRANSACTION_AMOUNT, CUST_CHG_AMOUNT]

    for row_name in money_rows:
        df = fix_money_row(df, row_name)
    
    main_entries = []
    offset_entries = []
    cogs_entries = []
    error_rows = []

    for _, row in df.iterrows():
        main_entry, offset_entry, cogs_entry, error = process_row(row)

        if error:
            error_rows.append(row)
            continue
        # Append results to corresponding lists if they exist
        main_entries.append(main_entry)
        offset_entries.append(offset_entry)
        cogs_entries.append(cogs_entry)
    
    file_content = ""
    for main, offset, cogs in zip(main_entries, cogs_entries, offset_entries):
        if main:
            file_content += make_row_fixed_width(main)
        if offset:
            file_content += make_row_fixed_width(offset)
        if cogs:
            file_content += make_row_fixed_width(cogs)

    with open(f"./samsung_uploads/G.GFEK100.TEXTIPTF_TRP_{file_name}_{_get_date_file_posting_format()}.txt", 'w+') as file:
        file.write(file_content)
    
    do_prod_stuff(file_path, file_content)
    main_entries = [d for d in main_entries if d] 

    send_email(file_name, main_entries)
    
    if error_rows:
        send_error_email(error_rows)


def do_prod_stuff(file_path, file_contents: str):   
    try:
        # Get the file's directory and filename
        dir_path, old_filename = os.path.split(file_path)
        new_filename = f"{os.path.splitext(old_filename)[0]}_completed.xlsx"  # Create new name
        new_file_path = os.path.join(dir_path, new_filename)

        # Rename the file
        os.rename(file_path, new_file_path)
        print(f"File renamed to: {new_file_path}")

        # Only write to sandbox iff the renaming went smoothly. 
        prod_file_name = f"{PROD_PATH}/PG.GFEK100.TEXTIPTF_TRP_{_get_date_file_posting_format()}.txt"
        with open(prod_file_name, "w+") as file:
            file.write(file_contents)
        print(f"WROTE FILE: {prod_file_name}")
    except Exception as e:
        exc_type, exc_value, exc_tb = sys.exc_info()
        formatted_tb = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
        error_message = f"Failed to rename file {file_path.name} to have _completed. Error: {str(e)}"
        print(error_message)
        MailUtil('email').send("samsung_gl issues bro!!", f"{file_path}\n YEah so like! We couldn't rename the file! Major Issues!!!\n\nStack Trace:\n{formatted_tb}") 
        raise OpenFileException()



def make_row_fixed_width(row: dict[str, Any]):
    operating_unit = justify(int(float(row[OPERATING_UNIT])), LEN_OPERATING_UNIT, '0')
    # operating_unit is the cost center. 
    
    division = justify(int(float(row[DIVISION])), LEN_DIVISION, '0')

    account = justify(row[ACCOUNT], LEN_ACCOUNT)

    source = justify(THIRD_PARTY_REPLACEMENT_PURCHASES, LEN_SOURCE)

    after_source_blurb = justify(' ', LEN_CATEGORY + LEN_YEAR + LEN_WEEK)

    cost = justify(row[COST], LEN_COST)                                             # Column Y or -R or P(.75)

    after_cost_blurb = justify(' ', LEN_CURRENCY_CODE + LEN_SELLING_VALUE + LEN_STATISTICAL_AMOUNT + LEN_STATISTICAL_CODE)

    reversal_flag = justify('N', LEN_REVERSAL_FLAG)

    after_reversal_flag_blurb = justify(' ', LEN_REVERSAL_YEAR + LEN_REVERSAL_WEEK + LEN_BACKTRAFFIC_FLAG + LEN_SYSTEM_SWITCH)

    description = justify("Samsung Replacement Purchase", LEN_DESCRIPTION)    # "Samsung Replacement Purchase"

    after_description_blurb = justify(' ', LEN_ENTRY_TYPE + LEN_RECORD_TYPE + LEN_REF_NBR_1)

    doc_nbr = justify(row[DOC_NBR], LEN_DOC_NBR)                                     # Column U
    doc_nbr = str(doc_nbr).replace('-', '').lower().replace("none", "").replace('.0', '')
    doc_nbr = justify(doc_nbr, LEN_DOC_NBR)
    if len(doc_nbr.strip()) != 12:
        raise Exception("WHAT? WHY IS THIS HAPPENING")

    ref_number_2 = justify(' ', LEN_REF_NBR_2)

    misc_1 = justify(str(row[REFERENCES_NUMBER]).strip().split(' ')[0], LEN_MISC_1)              #   Use Column N! (reference number)

    misc_2 = justify(str(row[MERCHANT_NAME]), LEN_MISC_2)

    misc_3 = justify(str(row[ACCOUNT_NUMBER]), LEN_MISC_3)

    after_misc_blurb = justify(' ', LEN_TO_FROM + LEN_DOC_DATE + LEN_EXP_CODE + LEN_EMP_NBR)

    row_date = datetime.strptime(str(row[DET_TRAN_DATE]), '%m/%d/%Y')

    # detail transaction date
    det_tran_date = justify(_get_det_tran_date(row_date), LEN_DET_TRAN_DATE)

    after_det_tran_blurb = justify(' ', LEN_ORIG_ENTRY + LEN_REP_FLG + LEN_ORU)

    # Sandy -> if this informs the posting period, it should be when you upload the file. 
    # gl transaction date informs the posting period. 
    gl_trxn_date = justify(row[POSTING_DATE], LEN_GL_TRXN_DATE)

    filler = justify(' ', LEN_FILLER)

    result = f"{operating_unit}{division}{account}{source}{after_source_blurb}{cost}{after_cost_blurb}{reversal_flag}{after_reversal_flag_blurb}{description}{after_description_blurb}{doc_nbr}{ref_number_2}{misc_1}{misc_2}{misc_3}{after_misc_blurb}{det_tran_date}{after_det_tran_blurb}{gl_trxn_date}{filler}\n"
    # result = f"{operating_unit}|{division}|{account}|{source}|{category}|{year}|{week}|{cost}|{currency_code}|{selling_value}|{statistical_amount}|{statistical_code}|{reversal_flag}|{reversal_year}|{reversal_week}|{backtraffic_flag}|{system_switch}|{description}|{entry_type}|{record_type}|{ref_number_1}|{doc_nbr}|{ref_number_2}|{misc_1}|{misc_2}|{misc_3}|{to_from}|{doc_date}|{exp_code}|{emp_nbr}|{det_tran_date}|{orig_entry}|{rep_flg}|{oru}|{gl_trxn_date}|{filler}\n"
    assert(len(result) == 301)
    return result


def _get_date_file_posting_format():
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def _get_det_tran_date(the_date: date):
    month = str(the_date.month).zfill(2)
    day = str(the_date.day).zfill(2)
    year = str(int(the_date.year))[-2:]
    return f"{month}{day}{year}"


def _get_gl_trxn_date(the_date: date):
    month = str(the_date.month).zfill(2)
    day = str(the_date.day).zfill(2)
    year = int(the_date.year)
    return f"{year}{month}{day}"


def send_email(file_name: str, df:List[str]):
    def highlight_cogs(s):
        return ['background-color: yellow' if v != 0 else '' for v in s]
    
    main_html_table = ""
    df_string_body = f"Using '{file_name}'\n"
    subject = f"{file_name} Auto Upload"

    main: pd.DataFrame = pd.DataFrame(df)

    if not main.empty:
        main: pd.DataFrame = main.style.apply(highlight_cogs, subset=[CUST_CHG_AMOUNT])
        main.columns = main.columns.str.lower()
        main_html_table += f"{main.to_html(index=False, escape=True)}\n"
        df_string_body += "The following 50905 entries have been submitted to the general ledger:\n" + main.to_string() + "\n"
    
    df_string_body += f"This is an automated email. Contact {os.getenv('email')} with any questions."
    if len(main_html_table) == 0:
        subject = f"{file_name} No Files Uploaded"
        html_body = f"""
            <html>
                <body>
                <h3>Using '{file_name}'</h3>
                <p>No records were submitted from {file_name}.</p>
                <p><small><em>This is an automated email. Contact {os.getenv('email')} with any questions.</em></small></p>
                </body>
            </html>
        """        
        df_string_body += "No records were submitted to the GL today."
    else:
        html_body = f"""
            <html>
                <body>
                <h3>Using '{file_name}''</h3>
                <p>The following table represents entries that have been submitted to the general ledger.</p>
                <p>Each row is submitted to account 50905, as well as an offset to account 63005.</p>
                <p>Highlighted rows have an additional COGS entry to the account 50195.
                {main_html_table}
                <p><small><em>This is an automated email. Contact {os.getenv('email')} with any questions.</em></small></p>
                </body>
            </html>
        """
    MailUtil(SAMSUNG_TO_LIST).send(subject, df_string_body, html_body)


def send_error_email(file_name: str, errored:List[str]):
    html_table = ""
    df_string_body = ""
    df: pd.DataFrame = pd.DataFrame(errored)
    subject = f"Errors in file {file_name}"
    if not df.empty:
        html_table += f"{df.to_html(index=False, escape=True)}\n"
        df_string_body += df.to_string()
    if len(html_table) == 0:
        print("WHAT IS GOING ON")
        MailUtil('email').send(f"{file_name} samsung_gl issues bro.", df_string_body, html_body) 
        return
    
    html_body = f"""
        <html>
            <body>
            <p>The following items have not been submitted to the GL. They had an error, likely with their doc numbers:</p>
            {html_table}
            <p><small><em>This is an automated email. Contact {os.getenv('email')} with any questions.</em></small></p>
            </body>
        </html>
    """
    MailUtil(SAMSUNG_TO_LIST).send(subject, df_string_body, html_body)


def get_file_paths() -> List[str]:
    dirs = os.listdir(PICKUP_LOCATION)
    undone_paths = []
    for dir in dirs:
        dir_lowered = dir.lower()
        if "_completed" not in dir_lowered:
            undone_paths.append(dir)
    return undone_paths
    return "Samsung Re-class Replacement Expense 10.06-10.17B 10-28-25.xlsx"


def main():
    in_files = get_file_paths()

    for file in in_files:
        if "~$" in file[:3]:            # if it's an open file, just ignore it. 
            continue
        try:     
            process_file(PICKUP_LOCATION / file)
        except OpenFileException:
            MailUtil('samsung_to_list_error').send(f"Could Not Process {file}", f"{PICKUP_LOCATION / file}\n Could not rename file, so aborted GL processing. Likely someone has the file open.\n\n\nStack Trace:\n{formatted_tb}") 
        except PermissionError:
            exc_type, exc_value, exc_tb = sys.exc_info()
            formatted_tb = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
            MailUtil('email').send("minor file permission error", f"{PICKUP_LOCATION / file}\n PermissionError, likely someone has the file open.\n\nStack Trace:\n{formatted_tb}") 
            continue
        except Exception:
            exc_type, exc_value, exc_tb = sys.exc_info()
            formatted_tb = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
            MailUtil('email').send("minor file errored", f"{PICKUP_LOCATION / file}\n Stack Trace:\n{formatted_tb}") 
    print(f"Attempted to process {len(in_files)} files")
    print("ALL DONE!")


if __name__ == "__main__":
    main()
