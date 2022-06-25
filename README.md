# An Analysis of Green Stocks

## Overview of Project

### Purpose

The objective of this project was to refactor code written in Microsoft Excel VBA that assembles performance metrics of green stocks into one easy to access location in an efficient manner.

## Results 

### Analysis of the Code

The first step in optimizing the performace of any code is to assess the codes original functionality. The original code ran well but had multiple for loops that created inefficiency: 
```  
 '4) Loop through tickers
   For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       '5) loop through rows in the data
       Worksheets(yearValue).Activate
       For j = 2 To RowCount
           '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
        End If
           '5b) get starting price for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                startingPrice = Cells(j, 6).Value
        End If
           '5c) get ending price for current ticker 
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                endingPrice = Cells(j, 6).Value
         End If

    Next j
```
Here is the runtime of the original code:

![This is an image](https://github.com/DanielBergan/stock-analysis/blob/main/Resources/Original_Code_Runtime_2017.png)
![This is an image](https://github.com/DanielBergan/stock-analysis/blob/main/Resources/Original_Code_Runtime_2018.png)

    
    And so the task became paring those down to simplify the process:
 
 ```  
 '2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerindex) = tickerVolumes(tickerindex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
         If Cells(i, 1).Value = tickers(tickerindex) And Cells(i - 1, 1).Value <> tickers(tickerindex) Then
             tickerStartingPrice(tickerindex) = Cells(i, 6).Value
         
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
           If Cells(i, 1).Value = tickers(tickerindex) And Cells(i + 1, 1).Value <> tickers(tickerindex) Then
            tickerEndingPrice(tickerindex) = Cells(i, 6).Value
        
        End If

            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerindex) And Cells(i + 1, 1).Value <> tickers(tickerindex) Then
            tickerindex = tickerindex + 1
        
        End If
        
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
       Cells(4 + i, 1).Value = tickers(i)
       Cells(4 + i, 2).Value = tickerVolumes(i)
       Cells(4 + i, 3).Value = tickerEndingPrice(i) / tickerStartingPrice(i) - 1
        
    Next i 
 ``` 
Here is the runtime of the refactored code:

![This is an image](https://github.com/DanielBergan/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)
![This is an image](https://github.com/DanielBergan/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

### Analysis of the Stocks

This selection of stocks preformed well in 2017 with only TERP finishing down for the year. 
![This is an image](https://github.com/DanielBergan/stock-analysis/blob/main/Resources/Stock_Performance_2017.png)

However, in 2018 the green stocks took a downturn with all but ENPH and RUN ending the year in the red.
![This is an image](https://github.com/DanielBergan/stock-analysis/blob/main/Resources/Stock_Perfromance_2018.png)

## Summary

### General Refactoring

Some of the advantages of refactoring code include:
  - Better code performace
  - Code is easier to understand
  - Reduces cost of future modifications

Dissadvantages include:
  - May include substaintial upfront cost
  - Can be very time consuming

### Refactoring Stock Analysis Code

Refactoring the original stock analysis code resulted in substantially reduced runtime as well as cleaner, easier to understand code. Considering it is a relatively simple script the disadvantage of being time consuming was not a major obstacle. Clean, well organized code that has adequate comments makes it easier to maintain and fine tune in the long term.
