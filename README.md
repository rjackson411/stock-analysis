# Stock Analysis with Excel VBA

## Overview of Project

### Purpose
The purpose of this project was to refactor Microsoft Excel VBA code to collect stock information in the year 2017 and 2018, and then determine whether or not the stocks are worth investing in. The goal was to increase the efficiency of the original code.

### Analysis
Before starting the refactoring process, I copied the code that was needed to create the chart headers, ticker array, and to activate the appropriate worksheet. The steps were then listed in the correct order to set the structure for the refactoring. See the below sample of the refactored code.


'1a) Create a ticker Index
tickerIndex = 0

'1b) Create three output arrays
Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single

''2a) Create a for loop to initialize the tickerVolumes to zero.
' If the next row’s ticker doesn’t match, increase the tickerIndex.
For i = 0 To 11
    tickerVolumes(i) = 0
    tickerStartingPrices(i) = 0
    tickerEndingPrices(i) = 0
Next i

''2b) Loop over all the rows in the spreadsheet.
For i = 2 To RowCount

    '3a) Increase volume for current ticker
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
    '3b) Check if the current row is the first row with the selected tickerIndex.
    'If  Then
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    End If
    
    '3c) check if the current row is the last row with the selected ticker
    'If  Then
     If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
     End If

        '3d Increase the tickerIndex.
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
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

### Stock Performance in 2017

![Stocks_2017](https://user-images.githubusercontent.com/105829106/172744579-270b2b3d-ff2c-4c4e-8a06-286f4f6461a4.PNG)



### Stock Performance in 2018


![Stocks_2018](https://user-images.githubusercontent.com/105829106/172744637-6141e722-3ba5-42cb-af8b-eacdc81c78fc.PNG)



### Summary

What are the advantages or disadvantages of refactoring code?

Refactoring makes the code cleaner, more organized, and run more efficiently. A big advantage of cleaner code is that it's easier to read for you and for others. The biggest disadvantage is time. It may take along time to completely refactor code and that could pose a problem if you're on a tight deadline.

How do these pros and cons apply to refactoring the original VBA script?

The biggest benefit from refactoring was an decrease in macro run time. The original code took approximately one second to run, whereas our refactored code takes significantly less time to run. Attached below are the screenshots that indicate the run time for our new analysis.



![VBA_Challenge_2017](https://user-images.githubusercontent.com/105829106/172744654-ebd425ed-355a-43b4-a75d-1b2a306a35a5.PNG)


![VBA_Challenge_2018](https://user-images.githubusercontent.com/105829106/172744673-f96d277e-0a47-438a-a919-87b6f25e3037.PNG)
