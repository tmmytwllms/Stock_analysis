# VBA Challenge
Repo for Module 2 work

## Purpose
The purpose of this analysis was to compare stocks using a refactored script to sort and format the data. This scrip took data from the 2017 and 2018 sheets as selected within the input box, and then returned the stock values based on each ticker, which were ordered in an array. After going through each row to each ticker, it returned the total volume and yearly returns for each stock in either 2017 or 2018. This script then formatted the title to update with the sheet selection, headers to be bold with border underneath, and formatted returns above 0 to be green and those below 0 to be red. This script also contains a timer, which starts on script execution and ends on completion, returning the amount of time it took to run the script based on the year selection. Lastly, the output sheet contains a button to run the analysis which promts the input box, and a button to clear the worksheet.

## Results

### Stock Performance
For the year 2017, nearly all stocks saw a positive return, except for the TERP stock, which saw a negative return of 7.2%. All but two of these stocks, being DQ and HASI, saw a trading volume of over 100 million. Conversely, 2018 saw a mostly negative return, except for the ENPH and RUN stocks, which saw positive returns over 80%. All but AY saw a trading volume over 100 million for the year.

### Execution Times
The script had similar execution times when it was ran for both the 2017 and 2018 data sets. The 2017 data set was one one-thousandth of a second faster. For both years, the script took less than a second to run.

### Comparison of Original and Refactored Scripts
The original script was much easier to write, as it allowed for the functions of a subroutine to be broken down into multiple macros. Additionally, this allowed for individual testing of each separate function we wanted to implement; However, this required much more coding in order to separate each subfunction for the calculation. The refactored script was more difficult to write, due to the fact that all functions were condensed into one subroutine. This required more precise placement of each function, as an incorrect order would return an inaccurate result or throw an error. The refactored script worked much better, as we were able to condense the input box, loops, conditions, and formatting all into one macro and run it all with the press of a button.

## Summary
