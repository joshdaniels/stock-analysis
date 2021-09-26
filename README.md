# stock-analysis

## Overview 

### Background
Or friend Steve is trying to help his parents determine wither or not to invest in DAQO an energy company. 

### Purpose
He has asked us to levverage our abilities in Excel and Visual Basic for Applications to help him determine if it is a performing stock to invest in or not.

## Results
As it turns out, DQ performed the highest out of the other stocks we analyzed in terms of return for 2017. The return was 199.4% Curiously, For that year it was trading the lowest total volume. in 2017 RUN  and ENPH had positive returns and healthy total volumes

In 2018 however,Their total volume skyrocketed from 35 to 107 and the return on that stock was -63% In 2018, RUN and ENPH had high returns and high total volume.
The total volume for RUN and ENPH about doubled from 2017 to 2018 and they still grew their return. Out of these years, These two seem to be the most profitable and stable. 

2017:

![2017 Table](https://github.com/joshdaniels/stock-analysis/blob/main/2017_table.png)


2018:

![2018 Table](https://github.com/joshdaniels/stock-analysis/blob/main/2018_table.png)

### About the Code

When we initially ran this we were not using arrays and we were running through every single cell in the data. This resulted in a slower completion time for both 2017 & 2018.

#### Original Code: 

``` 
Sub AllStocksAnalysis()
    
    Dim startTime As Single
    Dim endTime  As Single
    

    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    yearValue = InputBox("What year would you like to run the analysis on?")
    startTime = Timer
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
      'Create a header row
      Cells(3, 1).Value = "Ticker"
      Cells(3, 2).Value = "Total Daily Volume"
      Cells(3, 3).Value = "Return"
      
    
    'Initialize array of all tickers
    Dim tickers(11) As String
    
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
    
    'Initialize variables for starting price and ending price
    Dim startingPrice As Single
    Dim endingPrice As Single
    
    
    'Activate data worksheet
   Sheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    'Loop through tickers
    For i = 0 To 11

        ticker = tickers(i)
        totalvolume = 0
        
        'loop through rows in the data
        Sheets(yearValue).Activate
        For j = 2 To RowCount
        
           '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then

                totalvolume = totalvolume + Cells(j, 8).Value

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

        'Output data for current ticker
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalvolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

   Next i
    
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    
End Sub

Sub formatAllStocksAnalysisTable()

    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("A3:C3").Font.Color = RGB(104, 58, 133)
    Range("A3:C3").Font.Italic = True
    Range("A3:C3").Font.Size = 16
    Range("B4:B15").NumberFormat = "$#,##0.00"
    Range("C4:C15").NumberFormat = "0.00%"
    Columns("B").AutoFit
    Cells(4, 3).Interior.Color = vbGreen
    Cells(4, 3).Interior.Color = vbRed
    Cells(4, 3).Interior.Color = xlNone
    
    dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd
    
        If Cells(i, 3) > 0 Then
        
            'Color the cell green
            Cells(i, 3).Interior.Color = vbGreen
            
        ElseIf Cells(i, 3) < 0 Then
    
            'Color the cell red
            Cells(i, 3).Interior.Color = vbRed
            
        Else
        
            'Clear the cell color
            Cells(i, 3).Interior.Color = xlNone
    
        End If
        
        If Cells(i, 4) = 1 Then
        
            'Color the cell red
            Cells(i, 4).Interior.Color = vbRed
    
        End If
        
        If Cells(i, 5) = 1 Then
        
            'Color the cell red
            Cells(i, 5).Interior.Color = vbRed
    
        End If
    
    Next i
    
    
End Sub

Sub ClearWorksheet()

    Cells.Clear

End Sub

Sub Button1_Click()
    Call AllStocksAnalysis
    Call formatAllStocksAnalysisTable
End Sub

Sub Button1_Click1()
    Call DQAnalysis
    Call DQAnalysis1
End Sub




```

2017 Before Refactoring runtime:

![2017 Runtime](https://github.com/joshdaniels/stock-analysis/blob/main/2017_before.png)

2018 Before Refactoring runtime:

![2018 Runtime](https://github.com/joshdaniels/stock-analysis/blob/main/2018_before.png)


### Refactoring

By using arrays and nested loops, We were able to drastically reduce the amount of time it took to output the same results. 

#### Code after refactoring

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
    ' Create a ticker index
    tickerIndex = 0
    
    '3a) Initialize variables for starting price and ending price
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ' intialize total volume
    Dim tickerVolumes(12) As Long
       
       '3b) Activate data worksheet
       Worksheets(yearValue).Activate
       
       '3c) Get the number of rows to loop over
       RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
       '4) Loop through tickers and initialize totalvolume to 0
       For i = 0 To 11
           
           tickerVolumes(i) = 0
           
        Next i
        
        
           '5) loop through rows in the data
           For j = 2 To RowCount
           
           
               '5a) Get total volume for current ticker
            
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value

           '5b) get starting price for current ticker
           If Cells(j - 1, 1).Value <> tickers(tickerIndex) Then

               tickerStartingPrices(tickerIndex) = Cells(j, 6).Value

           End If

           '5c) get ending price for current ticker
           If Cells(j + 1, 1).Value <> tickers(tickerIndex) Then

               tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
               
                ' increase ticker index
                tickerIndex = tickerIndex + 1
                
           End If
       Next j
       
       
       '6) Output data for current ticker
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

Sub ClearWorksheet()

    Cells.Clear

End Sub

Sub Button1_Click()
    Call AllStocksAnalysisRefactored
End Sub
```

2017 After Refactoring runtime:

![2017 ARuntime](https://github.com/joshdaniels/stock-analysis/blob/main/2017_after.png)

2018 After Refactoring runtime:

![2018 ARuntime](https://github.com/joshdaniels/stock-analysis/blob/main/2018_after.png)

## Summary

