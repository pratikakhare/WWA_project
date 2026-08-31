import re
from functools import lru_cache
from io import BytesIO

import pandas as pd
from dateutil import parser
from openpyxl import Workbook

VALID_SUBMISSION_CATEGORY = {
    'FOB', 'OFR', 'PLC', 'SVC', 'ALL',
    'FOB/OFR', 'FOB/PLC', 'FOB/SVC',
    'OFR/PLC', 'OFR/SVC', 'PLC/SVC',
    'FOB/OFR/PLC', 'FOB/OFR/SVC',
    'FOB/PLC/SVC', 'OFR/PLC/SVC',
    'OFR (IPI)'
}

VALID_SUB_CATEGORY = {
    'ROUTING',
    'OFR SURCHARGES',
    'BASE CHARGES PER WM-MIN & BL',
    '3RD PARTY CFS CHARGES',
    'BASE CHARGES PER WM-MIN',
    'TRANSIT TIME',
    'CURRENCY',
    'CARRIER SCAC CODES',
    'ENTIRE SHEET',
    'SOLAS FEE',
    'FILING FEES',
    'ORIGIN-DESTINATION VIA',
    'COMPENSATION',
    'RATE BASIS',
    'STORAGE CHARGE AFTER FREE TIME',
    'BASE CHARGES PER BL',
    'POL/POD DETAILS',
    'OTHER SURCHARGES',
    'ALTERNATIVE PORTS',
    'HAZ CHARGES',
    'EXP. CLEARANCE SURCHARGE',
    'FREQUENCY',
    'BMSB',
    'CUTOFF DAYS',
    'DESCRIPTION',
    'HAZ ROUTING',
    'IPI',
    'OWN/CO-LOAD',
    'PRE-DEFINED RATES',
    'STORAGE FREE DAYS',
    'NON-STACKABLE CHARGE',
    'PALLETIZING PER PALLET'
}

VALID_ERROR_CATEGORY = {
    'DIFFERENT RATES FOR DUPLICATE LANES',
    'DUMMY RATES',
    'FILE SENT TO DIFFERENT EMAIL ID',
    'INCOMPLETE INFO - RATE DESCRIPTION MISSING',
    'INCOMPLETE INFO. - PARTIAL RATES',
    'INCORRECT ALT. UNCODE',
    'INCORRECT CURRENCY',
    'INCORRECT DETAILS TRANSMITTED FOR UPLOAD PURPOSE',
    'INCORRECT FILE FORMAT USED/TAMPERED',
    'INCORRECT HAZ DETAILS/RATES',
    'INCORRECT RATES/QUOTE/DETAILS',
    'INCORRECT SERVICE DETAILS',
    'INCREASED RATES',
    'LOCKED RATES OVERRIDE',
    'RECEIVED SUBMISSION POST DEADLINE',
    'MISMATCH IN SERVICE LANE BETWEEN MEMBERS',
    'MISSING DETAILS',
    'NO FEEDBACK',
    'PI NOT FOLLOWED',
    'RATES QUOTED FOR WM & TON',
    'RATES QUOTED WITH DIFFERENT VALIDITY',
    'ROUTING CHANGE-MISSING CONFIRMATION ON FOB/PLC',
    'TYPO ERROR',
    'RATES/DETAILS CORRECTION POST DEADLINE',
    'NEGATIVE RATES'
}

VALID_STATUS_OUTCOME = {
    'ADJUSTED AS PER STD PROCESS',
    'EXTENDED EXISTING RATES',
    'FIXED ERROR',
    'NO UPDATE PROCESSED (DECLINED BY OWNER)',
    'QUERY SENT TO MEMBERS FOR CONFIRMATION',
    'REMOVED',
    'UPDATED AS NS',
    'UPDATED REVISED RCVD RATES/DETAILS',
    'UPDATED ROUTING WITH NO CHANGE IN FOB/PLC',
    'UPDATED USING SDD',
    'NO FEEDBACK FROM ACCOUNT OWNER-NOT UPDATED'
}

VALID_HANDLERS = {
            'SR', 'AS', 'SM', 'SD', 'SK', 'KN', 'RP'
}

DATE_COLUMNS = [
    'RFQ RELEASE DATE',
    'DEADLINE',
    'RECEIVED',
    'FINAL SUBMISSION DATE'
]

MAPPING_YN = {
    'Y': 'YES',
    'YES': 'YES',
    'N': 'NO',
    'NO': 'NO'
}

REQUIRED_COLUMNS = {
    'GLOBAL/\nNAC',
    'NAMED ACCOUNT',
    'RFQ HANDLED BY',
    'WWA MEMBER EMAIL ID',
    'MONTH',
    'ENTRY COMPLETE YES / NO',
    'COMPLIANT (YES/NO)',
    'SUBMISSION CATEGORY',
    'SUBMISSION SUB-CATEGORY',
    'ERROR CATEGORY',
    'STATUS UPDATE/OUTCOME',
    'PENALTY REPORTED Y/N'
} | set(DATE_COLUMNS)

EMAIL_REGEX = re.compile(r'^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$')
WHITESPACE_RE = re.compile(r'\s+')
DIGIT_PREFIX_RE = re.compile(r'^\d+\s*[\.\)]\s*')
ITEM_SPLIT_RE = re.compile(r'(\d+)\.\s*(.*?)(?=\s*\d+\.|$)')

VALID_SUBMISSION_CATEGORY_PREFIXES = sorted(VALID_SUBMISSION_CATEGORY, key=len, reverse=True)
VALID_SUB_CATEGORY_PREFIXES = sorted(VALID_SUB_CATEGORY, key=len, reverse=True)
VALID_ERROR_CATEGORY_PREFIXES = sorted(VALID_ERROR_CATEGORY, key=len, reverse=True)
VALID_STATUS_OUTCOME_PREFIXES = sorted(VALID_STATUS_OUTCOME, key=len, reverse=True)


# =========================================================
# MASTER CLEANING FUNCTIONS
# =========================================================


@lru_cache(maxsize=16384)
def clean_email(email):
    if pd.isna(email) or not email:
        return ''

    email_str = str(email).strip()
    email_str = re.sub(r'[\n\r\t\xa0]', '', email_str)
    email_candidates = re.split(r'[;,|]+', email_str)

    if len([e for e in email_candidates if e.strip()]) != 1:
        return ''

    email = email_candidates[0].strip().strip('.,;').replace(' ', '')
    return email if EMAIL_REGEX.match(email) else ''


@lru_cache(maxsize=16384)
def parse_messy_date(x):
    if pd.isna(x) or str(x).strip() == '':
        return ''
    x = str(x)
    x = re.sub(r'IST', '', x, flags=re.IGNORECASE)
    x = re.sub(r'[()]', '', x)
    x = x.replace('//', '/')
    x = WHITESPACE_RE.sub(' ', x).strip()
    try:
        return parser.parse(x, fuzzy=True).strftime('%m/%d/%Y')
    except Exception:
        return ''


@lru_cache(maxsize=16384)
def clean_compliant(x):
    if pd.isna(x):
        return ''
    x = str(x).strip().upper()
    x = DIGIT_PREFIX_RE.sub('', x)
    x = WHITESPACE_RE.sub(' ', x)
    if 'REVISED' in x:
        return 'YES - REVISED RATES'
    if 'MAINTAIN' in x:
        return 'YES - MAINTAIN RATES'
    if x == 'NO' or ' NO ' in f' {x} ':
        return 'NO'
    return ''


multiline_cleaner_cache = {}

def multiline_cleaner(x, valid_set, valid_prefixes):
    if pd.isna(x):
        return ''
    x = str(x).upper().replace('\xa0', ' ')
    if x == '':
        return ''

    cache_key = (x, id(valid_set))
    if cache_key in multiline_cleaner_cache:
        return multiline_cleaner_cache[cache_key]

    x = WHITESPACE_RE.sub(' ', x).strip()
    if any(x.startswith(prefix) for prefix in valid_prefixes) and not re.match(r'^\d+\.', x):
        x = '1. ' + x

    x = re.sub(r'(\d+)\.', r' \1. ', x)
    x = WHITESPACE_RE.sub(' ', x).strip()

    matches = ITEM_SPLIT_RE.finditer(x)
    values = []
    for match in matches:
        val = match.group(2).strip()
        if val in valid_set:
            values.append(val)

    if not values and x in valid_set:
        values.append(x)

    if not values:
        result = ''
    else:
        result = '\n'.join(f'{i+1}. {val}' for i, val in enumerate(values))

    multiline_cleaner_cache[cache_key] = result
    return result


@lru_cache(maxsize=16384)
def normalize_month(value):
    month_map = {
        'JANUARY': 'JAN',
        'FEBRUARY': 'FEB',
        'MARCH': 'MAR',
        'APRIL': 'APR',
        'JUNE': 'JUN',
        'JULY': 'JUL',
        'AUGUST': 'AUG',
        'SEPTEMBER': 'SEP',
        'OCTOBER': 'OCT',
        'NOVEMBER': 'NOV',
        'DECEMBER': 'DEC',
        'NAN': ''
    }
    value = str(value).strip().upper()
    return month_map.get(value, value)


def clean_rfq_dataframe(df):
    for required_col in REQUIRED_COLUMNS:
        if required_col not in df.columns:
            df[required_col] = ''

    df['GLOBAL/\nNAC'] = df['GLOBAL/\nNAC'].astype(str).str.strip()
    df['NAMED ACCOUNT'] = df['NAMED ACCOUNT'].astype(str).str.strip()

    cond1 = ~df['GLOBAL/\nNAC'].isin(['GLOBAL', 'NAC'])
    blank_patterns = ['', '-', ' -', '- ', ' - ']
    cond2 = df['NAMED ACCOUNT'].isin(blank_patterns)
    df.loc[cond1 & cond2, 'GLOBAL/\nNAC'] = 'GLOBAL'
    df.loc[cond1 & ~cond2, 'GLOBAL/\nNAC'] = 'NAC'

    df.loc[~df['RFQ HANDLED BY'].isin(VALID_HANDLERS), 'RFQ HANDLED BY'] = ''
    df['WWA MEMBER EMAIL ID'] = df['WWA MEMBER EMAIL ID'].map(clean_email)

    for c in DATE_COLUMNS:
        df[c] = df[c].map(parse_messy_date)

    df['MONTH'] = df['MONTH'].map(normalize_month)
    df['ENTRY COMPLETE YES / NO'] = (
        df['ENTRY COMPLETE YES / NO']
        .astype(str)
        .str.strip()
        .str.upper()
        .map(lambda x: x if x in ['YES', 'NO'] else '')
    )
    df['COMPLIANT (YES/NO)'] = df['COMPLIANT (YES/NO)'].map(clean_compliant)
    df['SUBMISSION CATEGORY'] = df['SUBMISSION CATEGORY'].map(
        lambda x: multiline_cleaner(x, VALID_SUBMISSION_CATEGORY, VALID_SUBMISSION_CATEGORY_PREFIXES)
    )
    df['SUBMISSION SUB-CATEGORY'] = df['SUBMISSION SUB-CATEGORY'].map(
        lambda x: multiline_cleaner(x, VALID_SUB_CATEGORY, VALID_SUB_CATEGORY_PREFIXES)
    )
    df['ERROR CATEGORY'] = df['ERROR CATEGORY'].map(
        lambda x: multiline_cleaner(x, VALID_ERROR_CATEGORY, VALID_ERROR_CATEGORY_PREFIXES)
    )
    df['STATUS UPDATE/OUTCOME'] = df['STATUS UPDATE/OUTCOME'].map(
        lambda x: multiline_cleaner(x, VALID_STATUS_OUTCOME, VALID_STATUS_OUTCOME_PREFIXES)
    )
    df['PENALTY REPORTED Y/N'] = (
        df['PENALTY REPORTED Y/N']
        .astype(str)
        .str.strip()
        .str.upper()
        .map(MAPPING_YN)
        .fillna('')
    )
    return df


def process_rfq_file(uploaded_file):
    input_bytes = uploaded_file.read()
    input_buffer = BytesIO(input_bytes)

    sheets = [
        'Aseem', 'Samuel','Sonali', 
        'Sunil','Sachin', 'Rohan', 
        'Krushna'
    ]
    columns = None

    xls = pd.ExcelFile(input_buffer, engine='openpyxl')
    missing_sheets = [sheet for sheet in sheets if sheet not in xls.sheet_names]
    if missing_sheets:
        raise ValueError(f'Missing expected sheet(s): {", ".join(missing_sheets)}')

    workbook = Workbook(write_only=True)
    worksheet = workbook.create_sheet(title='Cleaned RFQ')
    header_written = False

    file_name = getattr(uploaded_file, 'name', '')
    if file_name and not file_name.lower().endswith(('.xlsx', '.xlsm', '.xltx', '.xltm')):
        raise ValueError('Please upload a valid Excel file (.xlsx, .xlsm, .xltx, .xltm).')

    for i, sheet in enumerate(sheets):
        if i == 0:
            df = xls.parse(sheet_name=sheet, header=7)
            columns = [str(c) if c is not None else '' for c in df.columns]
            for required_col in REQUIRED_COLUMNS:
                if required_col not in columns:
                    columns.append(required_col)
                    df[required_col] = ''
        else:
            df = xls.parse(sheet_name=sheet, header=None, skiprows=8)
            if df.shape[1] < len(columns):
                for c in range(df.shape[1], len(columns)):
                    df[c] = None
            elif df.shape[1] > len(columns):
                df = df.iloc[:, :len(columns)]
            df.columns = columns
            for required_col in REQUIRED_COLUMNS:
                if required_col not in df.columns:
                    df[required_col] = ''

        df.dropna(how='all', inplace=True)
        if df.empty:
            continue

        df = clean_rfq_dataframe(df)

        if not header_written:
            worksheet.append(columns)
            header_written = True

        for row in df.itertuples(index=False, name=None):
            cleaned_row = [None if pd.isna(value) else value for value in row]
            worksheet.append(cleaned_row)

    output_buffer = BytesIO()
    workbook.save(output_buffer)
    return output_buffer.getvalue()
