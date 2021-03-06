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
    ' If the next row???s ticker doesn???t match, increase the tickerIndex.
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
'If the next row???s ticker doesn???t match, increase the tickerIndex.
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

Step 3d increases the tickerIndex if the next row???s ticker doesn???t match the previous row???s ticker.

```
'3d Increase the tickerIndex.
      If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
      	tickerIndex = tickerIndex + 1
End If

Next i
```

Lastly, Step 4 loops through the arrays (tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices) to output the ???Ticker,??? ???Total Daily Volume,??? and ???Return??? columns in the spreadsheet.

```
'4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
Worksheets("All Stocks Analysis").Activate
Cells(4 + i, 1).Value = tickers(i)
Cells(4 + i, 2).Value = tickerVolumes(i)
Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
Next i
```

The rest of the code, provided in the starter code, formats for easier readability.

After finalizing my code, I ran the stock analysis and confirmed that the outputs for 2017 and 2018 were the same as they were in this module. 

In 2017, nearly all of the stocks performed well. The return for every stock was a net positive, except for ???TERP???. Some stocks reach returns between 100%-200%. All of the net positive returns are good investment opportunities for Steve and his parents. As for the code comparison, the execution time for the 2017 data using the original script was 1.180 seconds while the execution time for the 2017 data using the refactored code was 0.234 seconds.

![2017 original script time](https://user-images.githubusercontent.com/90656004/138200273-d1d018e9-2368-4dd7-ac0a-02ac053c8ae6.PNG)

![VBA_Challenge_2017](https://user-images.githubusercontent.com/90656004/138200303-06f75245-edaa-4bad-ad00-75eb22cc32b2.png)

Overall, the majority of stocks did not perform well in 2018. All except for ???ENPH??? and ???RUN??? resulted in net negative returns. The 2017 and 2018 data both showed that the stocks ???ENPH??? and ???RUN??? are great candidates for investment by Steve and his parents. By running both the original and refactored codes, the run times were as expected: the original script execution ran longer (1.375 seconds) than the refactored code (0.25 seconds). 
	
![2018 original script time](https://user-images.githubusercontent.com/90656004/138200285-e9fb8174-0654-4593-b00a-c5c3acef93d3.PNG)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/90656004/138200315-0c660740-938e-4be7-b45a-1151d1d9ea1c.png)


## Summary

- What are the advantages or disadvantages of refactoring code?

Refactoring code helps to improve the organization and readability of an original script. Looking at the code again and restructuring it can also help to find any missed errors, like duplicate or missing functions.  Additionally, refactoring may reveal any trends or patterns that may not have been obvious before. However, refactoring can also have its disadvantages. It can affect the testing outcomes if it is not restructured properly. What once may have been a logical structure can become confusing very quickly. It can also take a significant amount of time to restructure and ensure that the code still produces the expected outcomes. Overall, refactoring as a practice is a good idea, but it is heavily dependent on the data and its intended use. 

- How do these pros and cons apply to refactoring the original VBA script?

By refactoring the original VBA script, we were able to get faster run times for both the 2017 and 2018 data, as seen in the screenshots above. It made the stock analysis easier for the client, but a great deal of time was spent just to save a fraction of a second for the client. The question is how often Steve intends to analyze his stock data, which would likely be once a year when he gets new data to add. The frequency of Steve using the data needs to be weighed against the time it took to refactor the original script. 
