# Homework2--VBA-Challenge
VBA challenge for module 2 of SMU Data Science Course



## Instructions:
Simply open the "Developer" tab on a macro-enabled workbook of excel
Paste in the code found in the "tickerpull.txt" file found in this repository
Make sure data is organized properly (see notes)
Save the sub, return to developer tab
Run macro tickerpull()
Wait for the code to run
Ensure all values look correct before presenting.

### Notes:
This is a fairly rudamentary Sub. Your data has to be orgainized very specifically
1. Make sure all of the following titles are arranged IN THIS ORDER in columns A-G
   "ticker,	date, open,	high,	low,	close,	vol"
3. Date MUST NOT be stored as a string. On Alphabetical Testing, it is stored as a string but on the actual stock data sheet it is not. This relies on a Max function to find the highest date, and will not work with a string. 
4. Data MUST be sorted, first by ticker then by date
