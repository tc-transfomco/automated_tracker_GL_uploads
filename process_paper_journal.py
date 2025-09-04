import math
import os
from typing import Tuple, Dict
import warnings

import pandas as pd
from datetime import datetime, date, timedelta
import json



# P:\Finance\Accounts Payable\User Systems Support\CHECK ISSUING\PaperChecks


FIXED_WIDTH_SOURCE = 'CHK'      
# WEEK IS YEAR
# GL unit header is Operating Unit

P_DRIVE = '//Uskihfil4/Public/Finance/Accounts Payable/User Systems Support/CHECK ISSUING/PaperChecks/'
SANDBOX_DESTINATION = "//SHSNGSTFSX/shs_boomi_vol/Test/GL_Files/300-byte"
CONFIG_PATH = "./config.json"
PROD_PATH = '//SHSNGPRFSX/shs_boomi_vol/GL_Files/300-byte'

RE_ISSUE = 'Void & Re-Issue'
HOME_SERVICES = 'Home Services'
STEVEN_WARD = 'Steven Ward WEBPOS'
ONLINE_CHECKS = 'Online Checks AdhocAP'

JOURNAL_REQUIRED_HEADER = 'Journal Required'
JOURNAL_DONE_HEADER = 'Journal Done?'
GL_UNIT_HEADER = 'GL unit to charge'
GL_ACCOUNT_HEADER = 'GL account to charge'
GL_DIV_TO_CHARGE_HEADER = 'GL div to charge'
AMOUNT_HEADER = 'Amount'
HOFFMAN = 'Hoffman Check requests'
NEEDS_INVESTIGATION = 'Needs To Be Investigated'
ACCOUNT_NUMBER_HEADER = 'Number'
COST_CENTER_HEADER = 'CostCenter'
PAYEE_NAME = 'Payee Name'
MEMO = 'Memo'
CHECK_NUMBER = 'Check Number'
DOC_NUMBER = 'Doc Number #'

CHECK_PROCESSED_DATE = 'Check Processed Date'


CSV_HEADER_OPERATING_UNIT = 'operatingUnit'
CSV_HEADER_DIVISION = 'division'
CSV_HEADER_COST = 'cost'
CSV_HEADER_ACCOUNT = 'account'
CSV_HEADER_DESCRIPTION = 'description'

CSV_HEADER_TRANSFER_DATE = 'detTranDate'
CSV_HEADER_POSTING_YEAR = 'postingYear'
CSV_HEADER_POSTING_PERIOD = 'postingPeriod'
CSV_HEADER_REVERSAL_FLAG = 'reversalFlag'
CSV_HEADER_EMAIL = 'emailAddress'
CSV_HEADER_DOC_NUMBER = "docNo"

# csv_header_conversions = {GL_UNIT_HEADER: CSV_HEADERS[0],
#                           GL_DIV_TO_CHARGE_HEADER: CSV_HEADERS[1],
#                           AMOUNT_HEADER: CSV_HEADERS[2],
#                           GL_ACCOUNT_HEADER: CSV_HEADERS[3],
#                           CHECK_NUMBER_AND_MEMO: CSV_HEADERS[4],
#                           DOC_NUMBER: CSV_HEADER_DOC_NUMBER,
#                           }

CSV_HEADERS = [CSV_HEADER_OPERATING_UNIT, CSV_HEADER_DIVISION, CSV_HEADER_COST, CSV_HEADER_ACCOUNT,
               CSV_HEADER_DESCRIPTION, 'docDate', CSV_HEADER_DOC_NUMBER, 'refNbr1', 'refNbr2', 'misc1', 
               'misc2', 'misc3', 'toFrom', CSV_HEADER_TRANSFER_DATE, 'altLegalEntityFlag',
               'altLegalEntity', CSV_HEADER_POSTING_YEAR, CSV_HEADER_POSTING_PERIOD, 
               CSV_HEADER_REVERSAL_FLAG, CSV_HEADER_EMAIL]
CSV_NULL_VALUE = None

CHECK_NUMBER_AND_MEMO = 'Check number and Memo'

CHECKS_CUT_HEADERS = ['Check Type', JOURNAL_REQUIRED_HEADER, JOURNAL_DONE_HEADER, "File Date", CHECK_PROCESSED_DATE,
                      CHECK_NUMBER, PAYEE_NAME, MEMO, 'Reference', AMOUNT_HEADER, 'Address', 'Address 2', 'City',
                      'State', 'Zip', 'Country', 'TaxID / SSN', 'Method ID', 'Postage Type', 'Bank Code',
                      'Overnight Y\\N', GL_ACCOUNT_HEADER, GL_UNIT_HEADER, GL_DIV_TO_CHARGE_HEADER, DOC_NUMBER]

RE_ISSUE_HEADERS = ['Check Type', JOURNAL_REQUIRED_HEADER, JOURNAL_DONE_HEADER, 'Check Received Date',
                    CHECK_PROCESSED_DATE, CHECK_NUMBER, PAYEE_NAME, MEMO, 'Reference', AMOUNT_HEADER, 'Address',
                    'Address 2', 'City', 'State', 'Zip  ', 'Country', 'TaxID / SSN', 'Method ID', 'Postage Type',
                    'Bank Code', 'Overnight Y\\N', GL_ACCOUNT_HEADER, GL_UNIT_HEADER, GL_DIV_TO_CHARGE_HEADER,
                    'Authorization #']

JOURNAL_REQUIRED_STR_13_15 =    'Yes - between PNC bank GL 10318 & JPM bank GL 10617'
JOURNAL_REQUIRED_STR_16_61_65 = 'Yes - between PNC bank GL 10336 & JPM bank GL 10617'
JOURNAL_REQUIRED_STR_14 =       'No'


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
# result = 5 + 4 + 5 + 3 + 4 + 4 + 2 + 17 + 3 + 17 + 17 + 3 + 1 + 4 + 2 + 1 + 2 + 30 + 3 + 2 + 10 + 15 + 10 + 20 + 20 + 20 + 6 + 8 + 3 + 7 + 6 + 15 + 1 + 6 + 8 + 16
# print(result)
# exit(0)
class NoSuchFolderException(Exception):
    def __init__(self, *args, **kwargs):
        self.super().__init__(args, kwargs)


def get_abbreviated_month_and_year():
    current_datetime = datetime.now()
    short_month_name = current_datetime.strftime('%b')
    two_digit_year = current_datetime.strftime('%y')

    return short_month_name, two_digit_year


def get_journals_required_but_not_done(dataframe):
    def journal_required_and_not_done(row):
        journal_required = str(row[JOURNAL_REQUIRED_HEADER]).lower().strip()
        journal_done = str(row[JOURNAL_DONE_HEADER]).lower().strip()
        return (journal_required == 'yes') and (journal_done == 'no')

    working_on_mask = dataframe.apply(journal_required_and_not_done, axis=1)
    return dataframe[working_on_mask], working_on_mask


def save_automated_check_journal_entries(dataframe_dict, month=None, year=None):
    return # todo, remove this.
    if month is None or year is None:
        month, year = get_abbreviated_month_and_year()
    writer = pd.ExcelWriter(f'outputs/Automated Check Journal Entries {month}{year[-2:]}.xlsx')

    for key in dataframe_dict.keys():
        dataframe_dict[key].to_excel(writer, sheet_name=key, index=False)
    writer.close()

def save_tracker(tracker, name):
    # if name != ".\AUTOMATED_TEST.xlsx":
    #     global CONTER
    #     name = f"My test_{CONTER}.xlsx"
    #     CONTER += 1
    writer = pd.ExcelWriter(name)

    for key in tracker.keys():
        tracker[key].to_excel(writer, sheet_name=key, index=False)

    writer.close()


def write_tracker_with_updates(tracker, masks):
    today = date.today()
    for key in masks.keys():
        if key in tracker:
            tracker[key].fillna('', inplace=True)
            tracker[key][JOURNAL_DONE_HEADER] = tracker[key][JOURNAL_DONE_HEADER].astype(str)
            if NEEDS_INVESTIGATION not in key:
                tracker[key].loc[masks[key], JOURNAL_DONE_HEADER] = f"AUTOMATED UPLOAD on {today}"
    # TODO remove this. 
    print(f"{P_DRIVE}AUG 2025 - Period 07/Checks Cut Aug'25(newest_testing_thomas).xlsx")
    print("##############################################")
    save_tracker(tracker, f"{P_DRIVE}AUG 2025 - Period 07/Checks Cut Aug'25(newest_testing_thomas).xlsx")


def find_bad_gl_accounts(df, valid_gl_accounts):
    def account_exists(row):
        return row[GL_ACCOUNT_HEADER] in valid_gl_accounts[ACCOUNT_NUMBER_HEADER].array

    def account_does_not_exist(row):
        return row[GL_ACCOUNT_HEADER] not in valid_gl_accounts[ACCOUNT_NUMBER_HEADER].array

    not_exists_mask = df.apply(account_does_not_exist, axis=1)
    exists_mask = df.apply(account_exists, axis=1)
    return exists_mask, not_exists_mask


def get_dataframes_from_journal(save=True):
    try:
        file_name, fiscal_month, fiscal_year = _get_latest_excel_spreadsheet()
        print(file_name)
    except NoSuchFolderException:
        print("Couldn't access the P drive!, or there were no folders for the current fiscal month!")
        file_name = "data/inp_file.xlsx"
    dfs = pd.read_excel(file_name, sheet_name=None, engine='openpyxl')
    for i, key in enumerate(dfs[HOME_SERVICES].keys()):
        assert(CHECKS_CUT_HEADERS[i] == key.strip())
    
    filtered = {}
    masks = {}

    for key in dfs.keys():
        if key not in [STEVEN_WARD, HOFFMAN, HOME_SERVICES, ONLINE_CHECKS]:
            continue
        
        dfs[key][JOURNAL_DONE_HEADER] = dfs[key][JOURNAL_DONE_HEADER].fillna('')
        dfs[key][JOURNAL_DONE_HEADER] = dfs[key][JOURNAL_DONE_HEADER].astype(str)

        filtered[key], masks[key] = get_journals_required_but_not_done(dfs[key])
        filtered[key] = filtered[key].fillna('')

    if save:
        save_automated_check_journal_entries(filtered, fiscal_month, fiscal_year)
    
    re_issue = dfs[RE_ISSUE]

    return dfs, file_name, filtered, masks, re_issue, fiscal_month, fiscal_year


def replace_cost_centers(dict_of_dfs, valid_cost_centers, replacements):
    def update_gl_units(row):
        correct_value = row[GL_UNIT_HEADER]
        if correct_value not in valid_cost_centers[COST_CENTER_HEADER].array:
            if int(correct_value) in replacements.keys():
                correct_value = replacements[int(correct_value)]
            else:
                correct_value = 8352
        return correct_value

    for key in dict_of_dfs.keys():
        updated = dict_of_dfs[key].apply(update_gl_units, axis=1)
        if not updated.empty:
            dict_of_dfs[key].loc[:, GL_UNIT_HEADER] = updated
    
    return dict_of_dfs


def process_checks_cut(cost_center_replacements, do_transfer):

    dfs, file_name, filtered, masks, re_issue, fiscal_month, fiscal_year = get_dataframes_from_journal(save=True)

    # End of step 4
    # Step 6

    re_issue[CHECK_NUMBER] = re_issue[CHECK_NUMBER].astype('Int64').astype(str)
    re_issue[MEMO] = re_issue[MEMO].astype(str)

    def update_journal_required_value(row):
        if row[MEMO][:2] == "14":
            return JOURNAL_REQUIRED_STR_14
        if row[CHECK_NUMBER][:2] == row[MEMO][:2]:
            return "Yes"
        replace_str = "Yes"
        if row[MEMO][:2] == "13" or row[MEMO][:2] == "15":
            replace_str = JOURNAL_REQUIRED_STR_13_15
        elif row[MEMO][:2] == "16" or row[MEMO][:2] == "61" or row[MEMO][:2] == "65":
            replace_str = JOURNAL_REQUIRED_STR_16_61_65

        return replace_str

    updated = re_issue.apply(update_journal_required_value, axis=1)
    if not updated.empty:
        re_issue[JOURNAL_REQUIRED_HEADER] = updated
    for key in dfs.keys():
        if "issue" in key.lower():
            dfs[key] = re_issue
            break
    # save_tracker(dfs, name=file_name)
    # Time to get the files that do need a journal update. We'll gather them here, and use them later.
    def journal_required_mask(row):
        if not row[JOURNAL_DONE_HEADER] == math.nan:
            return False
        if row[MEMO][:2] == "14":
            return False
        if row[CHECK_NUMBER][:2] == row[MEMO][:2]:
            return False

        return True
    masks[RE_ISSUE] = re_issue.apply(journal_required_mask, axis=1)
    
    re_issue_to_add_to_final_csv = re_issue[masks[RE_ISSUE]]

    # Step 9 Done.

    # we don't do 16 and 17, because they're dumb.

    # Step 18
    netsuite_info = _get_accounts_and_cost_center()
    accounts = netsuite_info["Accounts"]
    cost_centers = netsuite_info["Cost Center"]

    for key in filtered.keys():
        exists_mask, not_exists_mask = find_bad_gl_accounts(filtered[key], accounts)
        if filtered[key][not_exists_mask].size != 0:
            dfs[f"{NEEDS_INVESTIGATION} {key}"[:30]] = filtered[key][not_exists_mask]
            filtered[key] = filtered[key][exists_mask]
            # save_tracker(dfs, file_name)
    # Step 18 Done
    # Step 19
    cost_centers = netsuite_info["Cost Center"]
    filtered = replace_cost_centers(filtered, valid_cost_centers=cost_centers, replacements=cost_center_replacements)
    # Step 19 Done

    # # time for stupidity
    # # step 20

    # def no_update_gl_units(row):
    #     correct_value = row[GL_UNIT_HEADER]
    #     if correct_value not in cost_centers[COST_CENTER_HEADER].array:
    #         correct_value = pd.NA
    #     return correct_value

    # if do_steven_ward:
    #     the_same_garbage_in_column_gl_unit = filtered[STEVEN_WARD].apply(no_update_gl_units, axis=1)
    #     last_column = filtered[STEVEN_WARD].shape[1]
    #     filtered[STEVEN_WARD].insert(loc=last_column, column='Z', value=the_same_garbage_in_column_gl_unit)

    # # step 21. This one sucks so much. Why do I just filter in the "original file"?
    # #          and then just move on? what kind of steps are these?

    # def no_update_gl_account(row):
    #     correct_value = row[GL_ACCOUNT_HEADER]
    #     if correct_value not in accounts[ACCOUNT_NUMBER_HEADER].array:
    #         correct_value = pd.NA
    #     return correct_value

    # the_same_garbage_in_column_gl_unit = filtered[HOFFMAN].apply(no_update_gl_units, axis=1)
    # last_column = filtered[HOFFMAN].shape[1]
    # filtered[HOFFMAN].insert(loc=last_column, column='Z', value=the_same_garbage_in_column_gl_unit)

    # the_same_garbage_in_column_gl_account = filtered[HOFFMAN].apply(no_update_gl_account, axis=1)
    # last_column = filtered[HOFFMAN].shape[1]
    # filtered[HOFFMAN].insert(loc=last_column, column='AA', value=the_same_garbage_in_column_gl_account)

    # # enough of the stupidity

    # Step 22 and 23
    def construct_check_number_and_name(row):
        check_number = str(row[CHECK_NUMBER])
        check_number = ''.join(ch for ch in check_number if ch.isdigit())

        # Capital f string so that it looks even more like FUCK
        formatted_str = F"CK{check_number} {row[PAYEE_NAME]}"
        if len(formatted_str) > 30:
            formatted_str = formatted_str[:30]
        return formatted_str
    
    print("Compiling data...")

    for key in filtered.keys():
        check_number_and_name = filtered[key].apply(construct_check_number_and_name, axis=1)
        if not filtered[key].empty:
            filtered[key].insert(loc=0, column=CHECK_NUMBER_AND_MEMO, value=check_number_and_name)
    
    check_number_and_name = re_issue_to_add_to_final_csv.apply(construct_check_number_and_name, axis=1)
    if not re_issue_to_add_to_final_csv.empty:
        re_issue_to_add_to_final_csv.insert(loc=0, column=CHECK_NUMBER_AND_MEMO, value=check_number_and_name)

    # data_to_export = pd.DataFrame()
    # what we want: from HOMESERVICES, STEVEN WARD and HOFFMAN  ===> GL Unit    (column x)   IN   "Cost Center"
    #                                                                GL Div     (column y)   IN   "Sub Accts"
    #                                                                Amount     (column k)   IN   "Cost"
    #                                                                GL account (column w)   IN   "Acct"
    #                                                                payee name (column a)   IN   "Description
    #                                                                Doc Number (column ?)"  IN   "DocNo"       .. maybe, unsure about this one. 
    #                                                                the date                     in the date
    columns_to_keep = [GL_UNIT_HEADER, GL_DIV_TO_CHARGE_HEADER, AMOUNT_HEADER, GL_ACCOUNT_HEADER, CHECK_NUMBER_AND_MEMO, DOC_NUMBER, CHECK_PROCESSED_DATE, PAYEE_NAME]

    export_dataframes = []
    for key in filtered:
        if DOC_NUMBER not in filtered[key].columns:
            filtered[key][DOC_NUMBER] = None
        if not filtered[key].empty:
            export_dataframes.append(filtered[key][columns_to_keep])
    if not re_issue_to_add_to_final_csv.empty:
        export_dataframes.append(re_issue_to_add_to_final_csv[columns_to_keep])
    data_to_export = pd.concat(export_dataframes)

    csv_header_conversions = {GL_UNIT_HEADER: CSV_HEADERS[0],
                              GL_DIV_TO_CHARGE_HEADER: CSV_HEADERS[1],
                              AMOUNT_HEADER: CSV_HEADERS[2],
                              GL_ACCOUNT_HEADER: CSV_HEADERS[3],
                              CHECK_NUMBER_AND_MEMO: CSV_HEADERS[4],
                              DOC_NUMBER: CSV_HEADER_DOC_NUMBER,
                              }
    data_to_export = data_to_export.rename(columns=csv_header_conversions)

    # Columns A, B and D need to have 5 digits and leading zeroes.
    data_to_export[CSV_HEADERS[0]] = data_to_export[CSV_HEADERS[0]].astype(str).str.zfill(5)
    # data_to_export[CSV_HEADERS[1]] = data_to_export[CSV_HEADERS[1]].astype(str).str.zfill(5)
    data_to_export[CSV_HEADERS[3]] = data_to_export[CSV_HEADERS[3]].astype(str).str.zfill(5)

    def make_negative(row):
        return float(row[CSV_HEADERS[2]]) * -1

    reverse_data = data_to_export.copy()
    print("Balancing transactions...")
    reverse_data = reverse_data.assign(**{CSV_HEADERS[0]: "08500"})             # operatingUnit                (column A)
    reverse_data = reverse_data.assign(**{CSV_HEADERS[1]: "530"})               # division                     (column B)
    reverse_data[CSV_HEADERS[2]] = reverse_data.apply(make_negative, axis=1)    # copies of cost, but negative (column C)
    reverse_data = reverse_data.assign(**{CSV_HEADERS[3]: "10617"})             # account                      (column D)
    #                                                                           # Do nothing for description   (column E)
    #                                                                           # Do nothing for doc numbers   (column ?)

    data_to_export = pd.concat([data_to_export, reverse_data])

    print("Converting data into fixed width format...")
    fixed_lines = data_to_export.apply(make_row_fixed_width, axis=1)
    with open(f'fixed_width_result.txt', 'w+') as file:
        for line in fixed_lines:
            file.write(line)
    with open(f'{SANDBOX_DESTINATION}/fixed_width_result{_get_date_csv_format()}.txt', 'w+') as file:
        for line in fixed_lines:
            file.write(line)
    # data_to_export.to_csv("outputs/result.csv", index=False)
    # if do_transfer:
    #     data_to_export.to_csv(f"{SANDBOX_DESTINATION}/daily_drop-{date.today().day}_{fiscal_month}_{fiscal_year}.csv", index=False, header=False)
    print("DONE!")

    write_tracker_with_updates(dfs, masks)
    # TODO: put the csv into netsuite!

    # TODO: get some sort of response, hopefully a succesful response!
    #
    # TODO: Update the tracker with what you did!  So that the next day you don't do the same work!


def justify(input, amount, fillchar=' '):
    return str(input)[:amount].rjust(amount, fillchar)


def make_row_fixed_width(row):
    # GL unit header is Operating Unit
    operating_unit = justify(int(float(row[CSV_HEADER_OPERATING_UNIT])), LEN_OPERATING_UNIT, '0')
    division = justify(row[CSV_HEADER_DIVISION], LEN_DIVISION)

    account = justify(row[CSV_HEADER_ACCOUNT], LEN_ACCOUNT)

    source = justify(FIXED_WIDTH_SOURCE, LEN_SOURCE)

    category = justify(' ', LEN_CATEGORY)

    year = justify(' ', LEN_YEAR)

    week = justify(' ', LEN_WEEK)

    cost = justify(row[CSV_HEADER_COST], LEN_COST)

    currency_code = justify(' ', LEN_CURRENCY_CODE)

    selling_value = justify(' ', LEN_SELLING_VALUE)

    statistical_amount = justify(' ', LEN_STATISTICAL_AMOUNT)

    statistical_code = justify(' ', LEN_STATISTICAL_CODE)

    reversal_flag = justify('N', LEN_REVERSAL_FLAG)

    reversal_year = justify(' ', LEN_REVERSAL_YEAR)

    reversal_week = justify(' ', LEN_REVERSAL_WEEK)

    backtraffic_flag = justify(' ', LEN_BACKTRAFFIC_FLAG)

    system_switch = justify(' ', LEN_SYSTEM_SWITCH)

    description = justify(row[CSV_HEADER_DESCRIPTION], LEN_DESCRIPTION)

    entry_type = justify(' ', LEN_ENTRY_TYPE)

    record_type = justify(' ', LEN_RECORD_TYPE)

    ref_number_1 = justify(' ', LEN_REF_NBR_1)

    # TODO If there is a 50905 without a doc number, don't upload it. 
    # TODO: for adhoc, there is a different header so do something about that. 
    doc_nbr = justify(row[CSV_HEADER_DOC_NUMBER], LEN_DOC_NBR)
    doc_nbr = str(doc_nbr).replace('-', '').lower().replace("none", "")
    doc_nbr = justify(doc_nbr, LEN_DOC_NBR)
    if len(doc_nbr.strip()) == 1:
        doc_nbr = justify(' ', LEN_DOC_NBR)
        
    # NOTE: they should only be on 50905 GL account to charge. This is the account variable right now. 
    account_copy = account
    if int(str(account_copy.strip())) != 50905:
        doc_nbr = justify(' ', LEN_DOC_NBR)

    ref_number_2 = justify(' ', LEN_REF_NBR_2)

    misc_1 = justify(str(row[PAYEE_NAME]).strip().split(' ')[0], LEN_MISC_1)

    _, _, last_name = str(row[PAYEE_NAME]).strip().partition(' ')
    misc_2 = justify(last_name, LEN_MISC_2)

    misc_3 = justify(' ', LEN_MISC_3)

    to_from = justify(' ', LEN_TO_FROM)
    
    doc_date = justify(' ', LEN_DOC_DATE)

    exp_code = justify(' ', LEN_EXP_CODE)

    emp_nbr = justify(' ', LEN_EMP_NBR)
    row_date = datetime.strptime(str(row[CHECK_PROCESSED_DATE]), '%m%d%Y')

    # detail transaction date
    det_tran_date = justify(_get_det_tran_date(row_date), LEN_DET_TRAN_DATE)

    orig_entry = justify(' ', LEN_ORIG_ENTRY)

    rep_flg = justify(' ', LEN_REP_FLG)

    oru = justify(' ', LEN_ORU)

    # Sandy -> if this informs the posting period, it should be when you upload the file. 
    # gl transaction date informs the posting period. 
    gl_trxn_date = justify(_get_gl_trxn_date(date.today()), LEN_GL_TRXN_DATE)

    filler = justify(' ', LEN_FILLER)

    result = f"{operating_unit}{division}{account}{source}{category}{year}{week}{cost}{currency_code}{selling_value}{statistical_amount}{statistical_code}{reversal_flag}{reversal_year}{reversal_week}{backtraffic_flag}{system_switch}{description}{entry_type}{record_type}{ref_number_1}{doc_nbr}{ref_number_2}{misc_1}{misc_2}{misc_3}{to_from}{doc_date}{exp_code}{emp_nbr}{det_tran_date}{orig_entry}{rep_flg}{oru}{gl_trxn_date}{filler}\n"
    assert(len(result) == 301)
    return result



def _get_period():
    month = int(date.today().month)
    if month == 1:
        month = 13
    return month - 1


def _get_date_csv_format():
    today = date.today()
    month = str(today.month).zfill(2)
    day = str(today.day).zfill(2)
    year = int(today.year)
    return f"{month}{day}{year}"


def _get_det_tran_date(the_date):
    month = str(the_date.month).zfill(2)
    day = str(the_date.day).zfill(2)
    year = str(int(the_date.year))[-2:]
    return f"{month}{day}{year}"


def _get_gl_trxn_date(the_date):
    month = str(the_date.month).zfill(2)
    day = str(the_date.day).zfill(2)
    year = int(the_date.year)
    return f"{year}{month}{day}"


def _get_year():
    return int(date.today().year)


def _get_latest_excel_spreadsheet():
    print("Getting latest spreadsheet...")
    # range over the values 1, 0, -1. This will check next month first, then the current and finally the previous month. 
    for i in range(1, -2, -1):
        exists, dir, month, year = check_month_exists_in_p_drive(offset=i)
        if exists:
            return get_checks_cut_path(os.path.join(P_DRIVE, dir)), month, year

    raise NoSuchFolderException("Can not find a proper file for this month!")


def check_month_exists_in_p_drive(offset: int) -> Tuple[bool, str, str, str]:
    if offset <= 0:
        day = date.today()
        # this loop wont run if offset is 0. (so no change.)
        for _ in range(-offset):
            # Subtract a month
            day = day.replace(day=1)
            day = day - timedelta(days=3)
        abbv_month = day.strftime("%b").lower()
        year = str(day.year)
    if offset > 0:
        day = date.today()
        for _ in range(offset):
            # Add a month
            day = day.replace(day=28)
            day = day + timedelta(days=5)
        abbv_month = day.strftime("%b").lower()
        year = str(day.year)

    dirs = os.listdir(P_DRIVE)
    for dir in dirs:
        dir_lowered = dir.lower()
        if abbv_month in dir_lowered and year in dir_lowered:
            print(f"Using month {day.strftime("%B")}")
            return True, dir, abbv_month, year

    return False, None, None, None


def get_checks_cut_path(base_dir):
    dirs = os.listdir(base_dir)
    for dir in dirs:
        if not os.path.isfile(os.path.join(base_dir, dir)):
            continue
        dir_lowered = dir.lower()
        if "checks" in dir_lowered and "cut" in dir_lowered and ".xlsx" in dir_lowered and "v2" in dir_lowered:
            return os.path.join(base_dir, dir)


def _get_accounts_and_cost_center():
    try:
        return _access_database()
    except Exception:
        print("TODO: fix getting accounts and cost centers")
        print("Accessing all valid and accounts and cost centers...")
        return pd.read_excel("data/ExportData_CostCenters_Accounts.xlsx", sheet_name=None)
        pass

    return pd.read_excel("data/ExportData_CostCenters_Accounts.xlsx", sheet_name=None)


def _access_database():
    raise Exception("Not yet implemented")


def main():
    with open(CONFIG_PATH, 'r') as f:
        config = json.load(f)
     
    raw_dict: Dict[str, int] = config.get("cost_center_replacements", {})  
    cost_center_replacements: Dict[int, int] = {int(k): v for k, v in raw_dict.items()}
     
    do_transfer = False         # set to false if testing. 
    process_checks_cut(cost_center_replacements, do_transfer)


if __name__ == "__main__":
    
    warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
    main()
