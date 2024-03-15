# VBA-challenge
Module 2 challenge
The following is an explanation of the steps taken by me to do this assignment:
For the first part,I started by declaring all the variables that I'd need with their respective data types.
Then I created the column headers using the Range().Value command.
After that I put in the syntax to retrieve the total no. of rows in the worksheet.
I, then, assigned initial values to "tablerow" and "openingrow" which are variables I'll use to keep count. "tablerow" - to keep count of the row on the analysis table that I'm trying to create and "openingrow" to keep count of the first row of every new ticker value.
Then I used a "for" loop to iterate throught all the rows of data.
Inside the loop, I introduced an "If" condition to calculate all the values I need for the analysis
I used multiple "If"  statements to do conditional formatting.
Before ending the "If" condition, I reset my various counters.
For the next additional analysis, I again declared variables with their data types; created row & column headers; retrieved the total no. of rows of the data I gathered through my first VBA code; assigned initial values to the variables that I'll use to replicate the "COUNTIF" command in excel.
Using "for" loop to iterate and "If" command to conditionally select data, I was able to retrieve the necesaary data and also format it.
I tried these basic codes on a single sheet in the smaller excel file i.e. alphabetical_testing.xlsx and it worked fine after some debugging.
I then tried the "For each" function to apply my code to all the sheets in the workbook and that worked fine too. (Also made changes to some syntaxes)
Lastly, I successfully applied the codes to the main assignment workbook.
I really had a good time doing this assignment. The most challenging task for me was to retrieve the Opening price of a stock.
