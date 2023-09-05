# Bank-Account-Check-Tool
We received a list of companies interested in receiving funding, but a requirement for obtaining the funding was to have an account with a specific bank. Therefore, I created a program that searches whether a given company has an account with the bank we are looking for.
The program gets nip numbers from Excel, then gets their bank numbers from mf.gov.pl API and checks if any of them belongs to specific bank

# Execution
The script take a list of companies with their NIP number, then asks mf.gov.pl API about their bank account numbers, checks if one of them belong to specific bank and puts companies that have bank account in this specific bank to the list in .xlsx
