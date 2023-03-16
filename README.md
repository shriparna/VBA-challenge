# VBA-challenge

## Module 2: VBA-challenge activity
<hr/>

### Details of the program or pseudo logic

<hr>
- Program Pesudo Logic

- Start of the program
- First define all variables including for the second challenge (percentage calculation)
- Record the starting time 
- For each worksheet
    - First delete the columns from I to Q to erase the previous calculations (only visible if there is already calculation displayed from previous run)
    - Display the worksheet name (just to show which worksheet it is currently processing)
    - Intitalize the variables
    - Use TotalRowCount to store the number of total rows in that sheet
    - Print the column headers including for the second challenge 
    - Format the columns for width or % or scientific (+E)
    - Initialize the variables for the second challenge
    - For each row in the worksheet (starting from row 2 to TotalRowCount)
        - Get the open value only if its the first row for each ticker
        - add the total stock value
        - If current row ticker is different than the next row ticker then record the close value for each ticker
        - Calculate the yearly change as close value of the last row of the ticker - open value of the first row of the ticker
        - Percentage change as the yearly change divided by open value of the first row of the ticker (no need to mutliply by 100 as column is already formattted for %)
        - Print the ticker, yearly change, percent change and total stock value in the appropriate print row and column
        - If yearly change is less than zero then change the background color to red else to green
        - If percentage change is more than the previous one then record this as the new greatest percentage increase
        - If percentage is less than the previous one then record this as the new greatest percentage decrease
        - If total stock volume is greater than the previous one then record this as the new greatest total stock volume
    - This concludes for the each row
- This concludes for the each sheet
- Record the ending time
- Display the difference from end time to start time as the time taken (as one of the challenge asks us to verify it completed within 5 minutes)
- End of the program
