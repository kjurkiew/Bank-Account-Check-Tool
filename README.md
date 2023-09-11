# Bank-Account-Check-Tool
We received a list of companies interested in receiving funding, but a requirement for obtaining the funding was to have an account with a specific bank. Therefore, I created a program that searches whether a given company has an account with the bank we are looking for.
The program gets nip numbers from Excel, then gets their bank numbers from mf.gov.pl API and checks if any of them belongs to specific bank

# Prerequisites
Python 3.x
Required Python libraries:
  requests
  openpyxl

# Usage
Prepare an Excel file containing a list of NIPs in one column (e.g., "list_test.xlsx"). Make sure the NIPs start from the second row (the first row is typically reserved for headers).

Modify the script's configuration by opening the script file (main.py) in a text editor. You can change the following settings:

input_excel_filename: The name of your input Excel file.
bank_number: The bank number you want to search for (e.g., '2030' for BNP Paribas).
Run the script.
The script will process the Excel file, check the NIPs against the bank number, and create a new Excel file with the results. The new file will contain NIPs associated with the specified bank.

Check the generated Excel file (e.g., "wyniki_2_<last_row>.xlsx") to view the results.

# Example
Suppose you have an Excel file ("list_test.xlsx") with NIPs and want to find NIPs associated with BNP Paribas (bank number '2030'). After running the script, you will get a new Excel file with the relevant NIPs.
