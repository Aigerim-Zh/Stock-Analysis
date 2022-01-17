Sub AllStocksAnalysis()

    Dim startTime As Single
    Dim endTime As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

'setting the startTime variable equal to the Timer function
'It is important that the Timer starts only after the user entered the year
    startTime = Timer
'1. Format the output sheet on the "All Stocks Analysis" worksheet.

    Sheets("All Stocks Analysis").Activate
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Yearly Return"
    
 ' 2. Creating an array for tickers
    
    'First, we want to find out how many stocks we have by exporting the unique ticker names in the output sheet "All Stocks Analysis"
    'Before, we did this manually with a pivot table
    
    Sheets(yearValue).Activate
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row

    Range("A1:A" & rowEnd).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Sheets("Unique Tickers").Range("A1"), Unique:=True
    
    'Then, if look at column A in the "All Stocks Analysis" sheet, we can copy the list of the unique ticker names for 12 stocks in the dataset
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
    
 'Creating a ticker index

    Dim tickerIndex As Integer
        tickerIndex = 0

  'Creating arrays for Volumes, Starting and Ending prices, already knowing that we are going to have data for 12 stocks

    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

    For i = 0 To 11
         tickerVolumes(i) = 0
    Next i
    
    'Looping through rows in the stock data

     Sheets(yearValue).Activate

     For i = 2 To rowEnd
        If Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        End If
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If

        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        End If

         'Increase the tickerIndex if the next row’s ticker doesn’t match the previous row’s ticker.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerIndex = tickerIndex + 1
        End If

    Next i

    For i = 0 To 11
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    Next i

    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

Sub formatAllStocksAnalysis()
    'To make sure the correct worksheet is active
    Worksheets("All Stocks Analysis").Activate
    'Making the headers bold
    Range("A3:C3").Font.Bold = True
    'Creating a bottom border
    With Range("A3:C3").Borders(xlEdgeBottom)
        .LineStyle = xlContinuois
        .Weight = xlThin
    End With
  'Formatting the numbers
    Range("B4:B15").NumberFormat = "#,##"
    Range("C4:C15").NumberFormat = "0.00%"
    'Autofitting the total volume column
    Columns("B").AutoFit
    
    'Conditional formatting
    dataRowStart = 4
    dataRowEnd = 15
    
    For i = dataRowStart To dataRowEnd
        If Cells(i, 3) > 0 Then
            Cells(i, 3).Interior.Color = vbGreen
        ElseIf Cells(i, 3) < 0 Then
            Cells(i, 3).Interior.Color = vbRed
        Else
            'Clear the cell color
            Cells(i, 3).Interior.Color = xlNone
        End If
    Next i

End Sub

Sub ClearWorksheet()
    Worksheets("All Stocks Analysis").Activate
    Cells.Clear
End Sub
