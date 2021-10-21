# Stocks Analysis

## Overview of Project

### Purpose

Steve has asked for help with an analysis of his stock data. The first project in this module involved determining stock performance by analyzing financial data using Visual Basic for Applications (VBA). Now, he wants to expand the dataset to include the entire stock market over the last few years. In order to make the analysis applicable to the entirety of his stock data, this project will refactor the solution code to loop through all the stock data one time in order to collect the same information done in the initial analysis in this module.

## Results

The entire refactored code I created is shown below, but a detailed explanation of each step after the starter code is provided as well.

```
Sub AllStocksAnalysisRefactored()

    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
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
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
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
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
```

For Step 1a, the tickerIndex is set equal to zero before looping over the rows.

```
'1a) Create a ticker Index
tickerIndex = 0
```

Step 1b shows that arrays are created for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.

```
'1b) Create three output arrays
Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single
```

Step 2a creates loop to initialize the tickerVolumes to zero.

```
''2a) Create a for loop to initialize the tickerVolumes to zero.
'If the next row’s ticker doesn’t match, increase the tickerIndex.
For i = 0 To 11
tickerVolumes(i) = 0
tickerStartingPrices(i) = 0
tickerEndingPrices(i) = 0
Next i
```	

Step 2b creates a loop that will loop over all the rows in the spreadsheet.

```
 ''2b) Loop over all the rows in the spreadsheet.
  For i = 2 To RowCount
```

Inside the loop in Step 2b, Step 3a has a script that increases the current tickerVolumes (stock ticker volume) variable and adds the ticker volume for the current stock ticker.
It uses the tickerIndex variable as the index.

```
'3a) Increase volume for current ticker
tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
```

Step 3b is an if-then statement to check if the current row is the first row with the selected tickerIndex. It is, so the current starting price is assigned to the tickerStartingPrices variable.

```
'3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
tickerStartingPrices(tickerIndex) = Cells(i, 6).Value

End If
```

Step 3c is an if-then statement to check if the current row is the last row with the selected tickerIndex. It is, so the current closing price is assigned to  the tickerEndingPrices variable.

```
'3c) Check if the current row is the last row with the selected ticker
       If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
End If
```

Step 3d increases the tickerIndex if the next row’s ticker doesn’t match the previous row’s ticker.

```
'3d Increase the tickerIndex.
      If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
      	tickerIndex = tickerIndex + 1
End If

Next i
```

Lastly, Step 4 loops through the arrays (tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices) to output the “Ticker,” “Total Daily Volume,” and “Return” columns in the spreadsheet.

```
'4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
Worksheets("All Stocks Analysis").Activate
Cells(4 + i, 1).Value = tickers(i)
Cells(4 + i, 2).Value = tickerVolumes(i)
Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
Next i
```

The rest of the code formats for readability, which was provided in the starter code as well.


![VBA_Challenge_2017](https://user-images.githubusercontent.com/90656004/138190312-661a2cc7-8374-464d-b900-bc5b4097d742.png)


![VBA_Challenge_2018](https://user-images.githubusercontent.com/90656004/138190316-a0769e5b-6bd0-455e-94c2-450cf4a11c0f.png)



## Summary

- What are the advantages or disadvantages of refactoring code?


- How do these pros and cons apply to refactoring the original VBA script?

