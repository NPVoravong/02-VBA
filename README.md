# VBA Scripting

## Prompt
Using provided stock data Create a script that will loop through all the stocks for one year and output the following information.
* Ticker Symbol
* Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
* The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
* The total stock volume of the stock.
* Conditional Formatting for positive and negative yearly change.

## Process
1. **Active each sheet and process the data**   
This is done using a `For Loop` to cycle through and activate the sheets after all the data on that sheet has been processed.  

2. **Name four new columns using `Range`**  
  New Columns:
    * Ticker
    * Yearly Change
    * Percent Change
    * Total Stock Volume  
    
3. **Calculate Total Stock Volume**  
A `for loop` reads the data from the ticker symbols column and adds the values from the volume column. To keep only the values for that ticker an `If Then` statement is used to check the current ticker symbol against the next symbol. Once the two symbols are not equal the loop will terminate. After termination the symbol being processed gets written to the "Ticker Symbol" column and the aggregate volume gets written in the "Total Stock Volume" column. In order for the row counting to function properly the variable holding to sum of the current ticker needs to be reset to zero. Otherwise it will get the sum of all tickers.

4. **Calculate Yearly and Percent Change**  
To get the yearly change we need to set up a series of `nested for loops`. The for loops are keeping track of which ticker symbol is being processed, the value in the open column for the first instance of the ticker and the value in the close column for the last instance of the ticker. Like in the previous subroutine once the name of the tickers aren't equal the loop is broken and the script moves to the next thing. In this case it is using basic math to get the values for yearly and percent change then writing them to the specified row and column.

5. **Formatting**  
Conditional formatting is used to assign a color to the values in the "Percent Change" column. Positive values get green and negative values get red. The "Percent Change" column also gets formatted to read as percent instead of decimal. All of the new columns get the autofit option passed on them to ensure the data is easy to read.

## Results
<img src="/images/2014_Stock_Data_VBA.jpg" height="auto">
<img src="/images/2015_Stock_Data_VBA.jpg" height="auto">
<img src="/images/2016_Stock_Data_VBA.jpg" height="auto">






