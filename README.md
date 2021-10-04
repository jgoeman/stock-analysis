# Stock-Analysis and VBA Refactor
Peforming analysis on a small set of stock market data through the use of VBA and refactoring the VBA code to enable a higher level of efficiency.
## Overview of the Project
The client loved the previous work done on the stock data, and would like to be able to run this analysis for the entire stock market over the last few years. It has been requested that the previously created code be adjusted to run faster.
### Purpose
There are two main purposes for this project: 
- Review and Analyze the previous stock data
- Refactor the previously written VBA code so the analysis can be ran faster to allow more efficiency in larger data sets.

## Results

### Analysis of Stock information
2017 was a great year for the stocks provided in the data. Every stock besides TERP had at least a positive return. While just viewing 2017 it can be understood why the client's family invested heavily into the DQ stock as it had the greatest returns of 2017 at 199.4%. Unfortuneately most of the stocks had negative returns in 2018 with only ENPH and RUN showing positive returns. DQ also went from the best returns, of the provided stocks, in 2017 to the worst returns in 2018. The family of the client should look into diversifying there portfolio.   
### Analysis of Refactoring
Refactoring the data provided large gains in speed. The difference in the code between the original and the newly refactored code (below) is that there are no nested loops in the new code. In the original a loop was used to go through the ticker information and then a nested loop was used to loop through rows in all of the data. The refactored coded uses a variable called Tickerindex to provide data on which ticker the code is on. Three separate loops were also used in the new code:
- A loop to initialize the ticker volumes to 0
- A loop to go through each row on the spreadsheet
- A loop to go through the arrays and output the Ticker, Total Daily Volume, and the Return


```

    '1a) Create a ticker Index
    tickerindex = 0
        
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
        
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
        For i = 0 To 11
        tickerVolumes(i) = 0

        Next i
        
        
    ''2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
        
        '3a) Increase volume for current ticker
            tickerVolumes(tickerindex) = tickerVolumes(tickerindex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
            If Cells(i, 1).Value = tickers(tickerindex) And Cells(i - 1, 1) <> tickers(tickerindex) Then
            tickerStartingPrices(tickerindex) = Cells(i, 6).Value
              
            End If
        
            '3c) check if the current row is the last row with the selected ticker
            If Cells(i, 1).Value = tickers(tickerindex) And Cells(i + 1, 1).Value <> tickers(tickerindex) Then
            tickerEndingPrices(tickerindex) = Cells(i, 6)
            
            End If
        
            'If the next row�s ticker doesn�t match, increase the tickerIndex.
         
            If Cells(i + 1, 1).Value <> tickers(tickerindex) And Cells(i, 1).Value = tickers(tickerindex) Then
            
            '3d Increase the tickerIndex.
            tickerindex = tickerindex + 1
            
            End If
            
        Next I
```

## Summary

### Advantages and Disadvantages of refactoring code
Some advantages to refactoring code:
- The time to run the program is reduced
- In some instances the code may become easier to read
- Saved time when reviewing code in the future

A few disadvantages to refactoring code are:
- The time to refactor code could be immense
- Mistakes could happen while refactoring code requiring more debugging
### How the Previous Pro's and Con's apply to refactoring the original VBA Script
The main Con to refactoring the original VBA script took me several hours to do, and I made several mistakes to debug while doing it. The increased time to refactor seems to be the main pain point. The Pro's seem to greatly outweigh the cons though. When looking at the run times alone comparing the orignal code to the refactored code for 2017 it took 15 seconds on the orignal and .234 seconds on the refactored version. The 2018 information show similiar results with 1.49 seconds on the original and .242 seconds to run on the refactored code. This can be obviously seen as relative. With a data set as small as this a difference in seconds does not mean much, but if we were to add several thousand data points then the difference in time is easy to see.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/36859475/135808539-737e5f85-85bd-4f19-8bdc-e7189c655176.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/36859475/135808551-29889715-7740-45f8-ab38-1c0c6c9fd508.png)
