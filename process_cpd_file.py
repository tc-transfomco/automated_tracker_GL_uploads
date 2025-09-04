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


def justify(input, amount, fillchar=' '):
    return str(input)[:amount].rjust(amount, fillchar)


def make_row_fixed_width(row):
    # WEEK IS YEAR
    # GL unit header is Operating Unit
    operating_unit = justify(row[CSV_HEADER_OPERATING_UNIT], LEN_OPERATING_UNIT, '0')

    division = justify(row[CSV_HEADER_DIVISION], LEN_DIVISION)

    account = justify(row[CSV_HEADER_ACCOUNT], LEN_ACCOUNT)

    source = justify(FIXED_WIDTH_SOURCE, LEN_SOURCE)

    category = justify(' ', LEN_CATEGORY)

    year = justify(_get_year(), LEN_YEAR)

    week = justify(_get_period(), LEN_WEEK)

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

    doc_nbr = justify(row[CSV_HEADER_DOC_NUMBER], LEN_DOC_NBR)

    ref_number_2 = justify(' ', LEN_REF_NBR_2)

    misc_1 = justify(' ', LEN_MISC_1)

    misc_2 = justify(' ', LEN_MISC_2)

    misc_3 = justify(' ', LEN_MISC_3)

    to_from = justify(' ', LEN_TO_FROM)
    
    doc_date = justify(' ', LEN_DOC_DATE)

    exp_code = justify(' ', LEN_EXP_CODE)

    emp_nbr = justify(' ', LEN_EMP_NBR)

    det_tran_date = justify(' ', LEN_DET_TRAN_DATE)

    orig_entry = justify(' ', LEN_ORIG_ENTRY)

    rep_flg = justify(' ', LEN_REP_FLG)

    oru = justify(' ', LEN_ORU)

    gl_trxn_date = justify(' ', LEN_GL_TRXN_DATE)

    filler = justify(' ', LEN_FILLER)

    result = f"{operating_unit}{division}{account}{source}{category}{year}{week}{cost}{currency_code}{selling_value}{statistical_amount}{statistical_code}{reversal_flag}{reversal_year}{reversal_week}{backtraffic_flag}{system_switch}{description}{entry_type}{record_type}{ref_number_1}{doc_nbr}{ref_number_2}{misc_1}{misc_2}{misc_3}{to_from}{doc_date}{exp_code}{emp_nbr}{det_tran_date}{orig_entry}{rep_flg}{oru}{gl_trxn_date}{filler}\n"
    assert(len(result) == 301)
    return result