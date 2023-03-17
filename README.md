# Module 2: VBA-challenge

## Author: Shridhar Kamat.
<hr>

### Details of the program or pseudo logic
#### Shridhar Kamat
#### Date: 3/16/2023

Package Contents:

1. 2018 Screenshot: [Screenshot 2023-03-14 at 6.16.09 PM.png](https://github.com/shriparna/VBA-challenge/blob/main/Screenshot%202023-03-14%20at%206.16.09%20PM.png)
2. 2019 Screenshot: [Screenshot 2023-03-14 at 6.16.29 PM.png](https://github.com/shriparna/VBA-challenge/blob/main/Screenshot%202023-03-14%20at%206.16.29%20PM.png)
3. 2020 Screenshot: [Screenshot 2023-03-14 at 6.16.48 PM.png](https://github.com/shriparna/VBA-challenge/blob/main/Screenshot%202023-03-14%20at%206.16.48%20PM.png)
4. VBA Code: [VBA-challenge.bas](https://github.com/shriparna/VBA-challenge/blob/main/VBA-challenge.bas)
5. Documentation: [README.md](https://github.com/shriparna/VBA-challenge/blob/main/README.md)


<hr>
- Pseudo Logic  

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