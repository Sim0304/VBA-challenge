MONU-VIRT-DATA-PT-08-2023-U-LOLC-MTTH - VBA Challenge

## Description
The follow VBA script analises stock (ticker) data for a given year within the entire workbook. This is done by identifying all the unique tickers in any given sheet, calculating the yearly change in from the first opening price to the last closing price for that given year.The result would be a numerical difference and the percentage difference. The script also calculated the total volume of that given ticker for any given script. The script will also find the highest total volume, highest increase and decrease in yearly percentage change.

The script starts by defining data types to each variable used. Then creating summary tables for each sheet, detailing ticker values and their respective yearly change, yearly percent change and total volume, while also creating a table for the highest total volume, highest increase and decrease in yearly percentage change.

The script applies to all sheets in the workbook, and looping down through all the row. Starting by searching to the first unique ticker, assigning the first open price, then searching for the last closing price and setting the Open_Price and Close_Price as those values respectively associated with that ticker. Once those values are found, calculating Yearly_Change by minusing Open_Price from Close_Price, and then for non-zero results, outputting a percentage. Then calculating the Total_Volume by summing all volume values related to that ticker. Setting those values into their appropriate cells. Then resetting the total volume counter to zero to start the loops again.

Within the worksheet loop, it searches the table created from the ticker loop, for the highest total volume, highest increase and decrease in yearly percentage change then outputs the values in the relevant cells including the tickers.



