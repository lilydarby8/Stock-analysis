# All Stocks Analysis

## Overview of Project
The analysis preformed on the 2017 and 2018 Green Stocks data was to better equip Steve and his parents for investing in stocks. His parents were particularly interested in the "DQ" stocks. After running that analysis and getting a negative 0.626 return a need arose to put together an "All Stocks Analysis" for Steve and his parents to look at different stocks and make an informed financial decision. 

### Purpose
To accurately display the Ticker’s total daily volume and what return they had. In doing this Steve and his parents could quickly see the successful or unsuccessful tickers. 

## Results of Analysis
The analysis ran through all the tickers to add up each of their individual total daily volumes and calculate their return of stocks. Included in the code was formatting and conditional functions. Those functions displayed the data cleaner and clearer.  The challenge I faced was getting an overflow error when I ran the completed code. I had to make sure that excel knew where to take the data from when the "for loop" commenced. I had to specify before each cell to take the data from, " Worksheets(yearValue).".

**below are two images showing what happened when I didn’t specify the worksheet and how I added it to run properly without error
![This is an image](https://github.com/lilydarby8/Stock-analysis/issues/1#issuecomment-1161996477)
![This is an image](https://github.com/lilydarby8/Stock-analysis/issues/1#issuecomment-1161996615)


 **Also included are screenshots of the run time of the original code and the refactored code.
 ![This is an image](https://github.com/lilydarby8/Stock-analysis/issues/1#issuecomment-1161995402)
 ![This is an image](https://github.com/lilydarby8/Stock-analysis/issues/1#issuecomment-1161996010)
 ![This is an image](https://github.com/lilydarby8/Stock-analysis/issues/1#issuecomment-1161995656)
 ![This is an image](https://github.com/lilydarby8/Stock-analysis/issues/1#issuecomment-1161996141)


### Analysis Explained 
In the 2017 analysis of all stocks the ticker’s return rate fell into the positive category with only one ticker being negative. In comparison the 2018 analysis of all stocks did not do nearly as well. All but two tickers are in the negative.  

##Code Explained
   The original "AllStocksAnalysis" code ran and delivered the desired outcome, but was not sufficient, in that there were unnecessary lines of code it had to go through. In the "AllStocksAnalysis", all 3,012 rows were read 12 times; once for each ticker symbol. 

With the final code, "AllStocksAnalysisRefactored", three additional arrays were added, giving one array for each of the four columns in the "All Stocks Analysis" sheet. In adding the arrays and if-statements the rows of data were read one time while collecting the output data for each ticker symbol. This resulted in fewer lines of code needing to be read giving a faster output time.
### Advantages and Disadvantages of Refactoring Code 
An advantage to refactoring code is the code running faster which makes a big difference when you have a large amount of data because it contributes to fewer lines of code needing to be executed.
A disadvantage is the time-consuming processes of refactoring an already functioning code.
### Advantages and Disadvantages of the Original and Refactored Vba Script 
An advantage in the original code was that it was made up of simple and to the point functions. The disadvantage would be that because of its simplicity it required excel to work harder to retrieve the data needed. As in going through the data multiple times instead of one. 

An advantage to the refactored vba script apart from it running faster was that it didn’t have nested for loops so it made the code easier to read. A disadvantage to the refactored vba script was the time-consuming process of going through and adding code as well as the proper syntax for the code to still output the same data as the original code. 
