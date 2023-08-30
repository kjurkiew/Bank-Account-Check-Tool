# The program gets nip numbers from Excel, then gets their bank numbers and checks if any of them belongs to BNP

import requests
from openpyxl import load_workbook, Workbook

BASE_URL = 'https://wl-api.mf.gov.pl/api/search/nip/'

# Getting a bank account number
def get_bank_account_numbers(nip, date):
    url = f'{BASE_URL}{nip}?date={date}'
    try:
        response = requests.get(url)
        response.raise_for_status()
        data = response.json()
        account_numbers = data.get('result', {}).get('subject', {}).get('accountNumbers', [])
        return account_numbers
    except requests.exceptions.RequestException as e:
        print(f"Error fetching data for NIP {nip}: {e}")
        return []

# Checking whether any of the downloaded account numbers belong to BNP
def is_bnp_account(account_number, bank_code):
    return account_number.startswith(bank_code)

def main():
    workbook_filename = 'list_test.xlsx'
    date = '2020-01-24'
    bank_code = '2030' # BNP PARIBAS
    output_filename = f'wyniki_{date}.xlsx'

    try:
        workbook = load_workbook(filename=workbook_filename)
        sheet = workbook.active
        start_row = 2
        end_row = 10

        nips = [str(value[0]) for value in sheet.iter_cols(min_row=start_row,
                                                           max_row=end_row,
                                                           min_col=7,
                                                           max_col=7,
                                                           values_only=True)]

        lista_bnp = [nip for nip in nips if any(is_bnp_account(acc, bank_code) for acc in get_bank_account_numbers(nip, date))]

        if lista_bnp:
            result_workbook = Workbook()
            result_sheet = result_workbook.active
            for i, nip in enumerate(lista_bnp, start=1):
                result_sheet.cell(row=i, column=1).value = nip
            result_workbook.save(output_filename)
            print(f"Results saved to {output_filename}")
        else:
            print("No results found.")

    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
