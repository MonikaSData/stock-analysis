# stock-analysis
Performing analysis on stock data to uncover trends

## Overview of Project
1. Create interactive and user friendly Excel workbook that analyze stocks dataset (2017 and 2018) using VBA code to help answer questions about specific stocks performance. End user is able to chose which year should be analyzed - 2017 or 2018.
2. Measure run time of VBA code using nested loops and compare it to run time of refactored VBA code that ensures to loop through all the data one time in order to collect the same information as the original VBA code. The end goal is to determine whether refactoring made the VBA script run faster.

## Results


### Analysis Source File

The interactive Excel workbook can be viewed here [VBA Challenge](VBA_Challenge.xlsm)

### Run Time of VBA Code using Nested Loops

Code example:
   
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
        
        '5) Loop through rows in the data.
        Worksheets(yearValue).Activate
        For j = 2 To rowend
            '5a) Find the total volume for the current ticker.
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
            
            '5b) Find the starting price for the current ticker.
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1) = ticker Then
                startingPrice = Cells(j, 6).Value
            End If
            
            '5c) Find the ending price for the current ticker.
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1) = ticker Then
                endingPrice = Cells(j, 6).Value
            End If
            
        Next j

Analysing 2017 stocks data (original VBA code run time):

![2017Original](Resources/VBA_Challenge_Original_Code_2017.png)

Analysing 2018 stocks data (original VBA code run time):

![2018Original](Resources/VBA_Challenge_Original_Code_2018.png)

---
### Run Time of Refactored VBA Code

Example of refactored VBA code eliminating nested loops:

    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
        
    '2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
        ticker = tickers(tickerIndex)
        
        '3a) Increase volume for current ticker
          If Cells(i, 1).Value = ticker Then
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
          End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
         If Cells(i - 1, 1).Value <> ticker And Cells(i, 1) = ticker Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
           
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        If Cells(i + 1, 1).Value <> ticker And Cells(i, 1) = ticker Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
        
    Next i
    
    
Analyzing 2017 stocks data (refactored VBA code run time):

![2017](Resources/VBA_Challenge_2017.png)

Analyzing 2018 stocks data (refactored VBA code run time):

![2018](Resources/VBA_Challenge_2018.png)
  
---

## Summary

- What are the advantages and disadvantages of refactoring a code?

   *The advantages of refactoring are:*
   
       a. The code becomes more readable,less complex and easier to understand. 
       b. The process in most cases will uncover bad patterns in the code that can be eliminated. 
       c. Refactoring also leads to code/program that runs faster and is more adaptible. Debugging might become easier with refactored code.
   
   *The disadvantages of refactoring are:* 
   
      a. The process of refactoring can become time consuming and expensive. 
      b. If refactoring is not done right (there is not enough time to test the refactored code), the code can become broken which might lead to more bugs and issues during a release.

- What are the advantages and disadvantages of refactoring of the original and refactored VBA script?

  *The advantages of the refactored VBA stocks analysis script:*
  
      a. The script runs faster after refactoring (run time decreased from ~ 0.8 seconds to ~ 0.16 seconds)*
      b. The script is easier to understand*
      c. Using arrays and variables help the code to be more adaptable and reusable*
      
  *The disadvantages of the refactored VBA stocks analysis script:*
  
      a. The process to refactor the VBA code was time consuming*
      b. About 70% of the code had to be re-written and it might have been be easier and faster to write the code from scratch*
   
      
  
