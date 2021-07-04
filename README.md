# Module-2-Challenge-StockAnalysis

## Overview of Project
### Background
The parents of a friends are looking for investing in Green Energy as it is becoming more relevant and important, they believe will provide good return of investment however the company they are investing their money is rather based on gut feeling than data driven analysis which might be too risky specially for a growing and immature industry. 
The following project and purpose will help to identify the facts for making investments and informed decision. 

### Purpose

This project facilitates an annual volume and return analysis for the stock market to identify best performers stocks for investment at computing all available tickers from a given dataset to summarize volumes and visually indicate positive returns vs negative returns (losses).
Additionally the run time will be a factor to support future analysis as larger datasets may take several seconds or minutes to execute and the possibility to time out or  hang the workbook if the program does not perform well; therefore an optimized process is required for stock decision makers. 

## Analysis and Results

### Stock Performance (General Insight) 

According to the Analysis results 2017 was a better year for the Green Energy industry, the majority of the stocks had a positive and significant return of investment.
ENPH was the best performer taking in consideration both years: 2017 & 2018; DQ had a great 2017 year with high returns but wasn't the case for 2018 (Hope that Steve's parents invested just that year before their due diligence request)

![Refactored VBA Code 2017 Results & Run Time](https://github.com/Mejikano/Module-2-Challenge-StockAnalysis/blob/main/Resources/VBA_Challenge_2017.PNG)
![Refactored VBA Code 2018 Results & Run Time](https://github.com/Mejikano/Module-2-Challenge-StockAnalysis/blob/main/Resources/VBA_Challenge_2018.PNG)

### Code Performance
Comparing the refactored code performance versus the original code is evident that the times of iterating throughout the dataset records plays a key role on performance the less you iterate the most computing cycles are saved and performs better!

Following images show the original code run times for both 2017 and 2018. 

![Original VBA Code 2017 Results & Run Time](https://github.com/Mejikano/Module-2-Challenge-StockAnalysis/blob/main/Resources/Original_AllStockCode_2017.PNG)
![Original VBA Code 2018 Results & Run Time](https://github.com/Mejikano/Module-2-Challenge-StockAnalysis/blob/main/Resources/Original_AllStockCode_2018.PNG)

Why performance is so different?

-Because the original code reads all dataset entries/rows per each ticker being analyzed so the program **computes** the number of tickers (Tickers loop: For j = 0 To 11) by the number of data rows (rows loop:  For i = rowStart To rowEnd) **times**

```
'4) Loop through tickers
    For j = 0 To 11
    
        '
        Worksheets(yearValue).Activate
        totalVolume = 0
        
        'This loop goes through all the records/row to compute the calculations
        '5) loop through rows in the data
        For i = rowStart To rowEnd
        
            '5a) Get total volume for current ticker
            'increase totalVolume
            If Cells(i, 1).Value = tickers(j) Then
                totalVolume = totalVolume + Cells(i, 8).Value
            End If
            
            '5b) get starting price for current ticker
            If Cells(i - 1, 1).Value <> tickers(j) And Cells(i, 1).Value = tickers(j) Then
    
                startingPrice = Cells(i, 6).Value
    
            End If
            
            '5c) get ending price for current ticker
            If Cells(i + 1, 1).Value <> tickers(j) And Cells(i, 1).Value = tickers(j) Then
    
                endingPrice = Cells(i, 6).Value
    
            End If

        Next i
        
        '6) Output data for current ticker
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + j, 1).Value = tickers(j)
        Cells(4 + j, 2).Value = totalVolume
        Cells(4 + j, 3).Value = (endingPrice / startingPrice) - 1        
    Next j
	
```
-The refactor code read the rows of data once and store each ticker values in their corresponding arrays for Volumes, Starting Prices and End Prices.

```
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
            If Cells(i, 1).Value = tickers(tickerIndex) Then
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            End If
            
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
    
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    
            End If
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
        'If  Then
            
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
    
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                '3d Increase the tickerIndex.
                tickerIndex = tickerIndex + 1
                
            End If

    
        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1
        
    Next i

```



## Summary

### Original Code
The original code is inefficient as it iterates the full data set by the number of tickers; this is the major concern for the overall program and might cause performance issues with larger set of data.
One very quick refactor improvement could be to exit the loop when the end price for a given ticker is identified "break for loop" so the rest of the rows/entries are not iterated, and it would improve performance and reduce the number of times that the data set is read 
However I would argue too that this is not an elegant way to fix the performance concern (still code smell) and neither a better option for performing better as it the following refactoring was.   

### Refactored Code

**The way of reading the rows of data is the main advantage**
This refactored program significantly improves performance by reading the full dataset only once which definitely addresses the original code performance problem
However, I would list the following issues to be considered for future refactor looking for continuous  improvements:

**General disadvantages for both: Original Code and Refactored Code**

	1.Programs assume that the dataset contains rows/entries coming in order by ticker and descending dates which is unlike to happen for other years. This might cause:
		- Wrong start and end price or unhandled exceptions
		Proposed solution: this could be easily solved if the macro sorts by ticker (Column A) and date (Column B) at the beginning  of the program.
	2.The tickers are hardcoded; therefore the arrays contain magic numbers when tickers might be different for other years analysis i.e. New Green Energy companies may start trading on the stock market
		- This would cause wrong analysis excluding new companies 
		Proposed solution: Dynamically create an array and determine the value by a DISTINCT formula/method/function to identify the number of tickers on the dataset and their names.
	
			i.e. Below snippet demonstrates how arrays could be refactored to be dynamics
			```
			'Good dynamically VBA arrays reference @Stackoverflow: https://stackoverflow.com/questions/4326678/dynamically-dimensioning-a-vba-array
			    Dim tickerVolumes() As Long
				Dim tickerStartingPrices() As Single
				Dim tickerEndingPrices() As Single
				
			'Steps to figure out the distinct tickers counts and names then 
				ReDim tickerVolumes(numoftickers)
				ReDim tickerStartingPrices(numoftickers)
				ReDim tickerEndingPrices(numoftickers)
			```

	3.The program does not handle errors i.e.
		- Division by 0 will cause an exception/error - Start price could be 0, specially for new companies trading on the stock market
		- Inputting sheet names with different format and/or non-existing sheets 
	
		Proposed solution: 
			-Division by 0 can be handled catching errors or a similar approach as of the IFERROR formula learned before.
			-Write a code for checking whether given input sheet name exist before executing the next code steps; display an error message if it does not exist
	


### Analysis Workbook Reference
Below link has the Excel workbook used for this analysis with VBA macros including the Module 2 activities and Skill Drills (The Checkerboard one was interesting)

[Wall Street Stock Analysis workbook](https://github.com/Mejikano/Module-2-Challenge-StockAnalysis/blob/main/VBA_Challenge.xlsm)

