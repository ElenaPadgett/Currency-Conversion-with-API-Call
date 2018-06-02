**Finance Project with API**: 
The program retrieves currency conversion rates with fixer.io API and performs calculations in Excel.

**Motivation**: 
I wrote this piece of code to automate some commonly performed tasks in Excel.

1) Opens a given Excel file with transactions data (account, security name, security symbol, date, buy/sell, quantity, currency, local price)
2) Checks for exceptions, i.e. if 'Quantity' is an integer
3) Calculates the Total Amount of transaction in local currency
4) Pulls conversion rates using fixer.io API call.
5) Calculates Total Amount in EUR (using the retrieved conversion rates)
6) Calculates the broker commission based on the Account info
7) Creates two output files: Data_copy.xlsx with all the calculated fields populated and SummaryFile.csv.
