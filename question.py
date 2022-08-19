# The program gets nip numbers from Excel, then gets their bank numbers and checks if any of them belongs to BNP

import requests
from openpyxl import load_workbook, Workbook

workbooky = load_workbook(filename='list_test.xlsx')
sheet = workbooky.active
a = 2 # first verse
b = 10 # last verse

for value in sheet.iter_cols(min_row=a,
                             max_row=b,
                            min_col=7,
                            max_col=7,
                        values_only=True):
    nips = value

nipk = tuple(map(str, nips))

bank = '2030' # BNP PARIBAS
nipy = nipk
data = '2020-01-24'
lista_bnp = []

# Getting a bank account number
def get_bank_number(nip, data):
    r = requests.get('https://wl-api.mf.gov.pl/api/search/nip/'+nip+'?date='+data)
    r = r.json()
    r = r ['result']
    acc = r ['subject']
    if acc == None:
        acc = ['75249000050000453073876066']
    else:
        acc = acc ['accountNumbers']
    return acc

# Checking whether any of the downloaded account numbers belong to BNP
def is_bnp(nip, acc, list_bnp):
    for x in acc:
        if x[2:6] == bank:
            list_bnp.append(nip)
    return list_bnp

for nip in nipy:

    acc = get_bank_number(nip, data)
    is_bnp(nip, acc, lista_bnp)

print(lista_bnp)

book = Workbook()
sheet_2 = book.active
r = 1
for x in lista_bnp:
    sheet_2.cell(row=r, column=1).value = x
    r += 1

name = 'wyniki' + str(a) + '_' + str(b)
book.save(name + ".xlsx")
