# The program gets nip numbers from Excel, then gets their bank numbers and checks if any of them belongs to BNP

import requests
from openpyxl import load_workbook, Workbook
from datetime import date

# Getting a bank account number
def get_bank_number(nip, date):
    api_url = f'https://wl-api.mf.gov.pl/api/search/nip/{nip}?date={date}'

    try:
        response = requests.get(api_url)
        response.raise_for_status()  # Raise an exception for HTTP errors (4xx and 5xx)
        data = response.json()

        subject = data.get('result', {}).get('subject')

        if subject is None:
            account_numbers = ['75249000050000453073876066']
        else:
            account_numbers = subject.get('accountNumbers', [])

        return account_numbers
    except requests.exceptions.RequestException as e:
        # Handle any HTTP request errors here
        print(f"Error making the request: {e}")
    except ValueError as e:
        # Handle JSON decoding errors here
        print(f"Error decoding JSON response: {e}")

    return []  # Return an empty list if there was an error

# Checking whether any of the downloaded account numbers belong to BNP
def is_bnp(nip, account_number, bnp_account_numbers, bank_number):
    for account_number_prefix in account_number:
        if account_number_prefix[2:6] == bank_number:
            bnp_account_numbers.append(nip)
    return bnp_account_numbers

def get_maximum_rows(*, sheet_object):
    rows = 0
    for max_row, row in enumerate(sheet_object, 1):
        if not all(col.value is None for col in row):
            rows += 1
    return rows

def process_excel_file(filename, bank_number):
    workbook = load_workbook(filename=filename)
    sheet = workbook.active
    first_row = 2
    last_row =  get_maximum_rows(sheet_object=sheet)

    nips = []
    for value in sheet.iter_cols(min_row=first_row, max_row=last_row, min_col=7, max_col=7, values_only=True):
        nips.extend(map(str, value))

    date_today = date.today()
    lista_bnp = []

    for nip in nips:
        acc = get_bank_number(nip, date_today)
        is_bnp(nip, acc, lista_bnp, bank_number)
    print(lista_bnp)

    book = Workbook()
    sheet_2 = book.active
    r = 1
    for x in lista_bnp:
        sheet_2.cell(row=r, column=1).value = x
        r += 1

    output_filename = f'wyniki_{first_row}_{last_row}.xlsx'
    book.save(output_filename)

if __name__ == "__main__":
    input_excel_filename = 'list_test.xlsx'
    bank_number = '2030'  # BNP PARIBAS
    process_excel_file(input_excel_filename, bank_number)
