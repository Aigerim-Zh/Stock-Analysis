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
        Cells(3, 3).Value = "Return"
    
'2. Initialize an array of all tickers.

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

'3. Prepare for the analysis of tickers.
 '3a). Initialize variables for the starting price and ending price.
 
    Dim startingPrice As String
    Dim endingPrice As String

 '3b). Activate the worksheet from where we will pull the stock data
    Sheets(yearValue).Activate

 '3c). Find the number of rows to loop over.
    
     rowEnd = Cells(Rows.Count, "A").End(xlUp).Row

'4. Loop through the tickers.

    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0

'5. Loop through rows in the data.
        Sheets(yearValue).Activate
        
            For j = 2 To rowEnd
               '5a).Find the total volume for the current ticker.
               If Cells(j, 1).Value = ticker Then
                   totalVolume = totalVolume + Cells(j, 8).Value
               End If
              '5b).Find the starting price for the current ticker.
               If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
                   startingPrice = Cells(j, 6).Value
                 End If
    
               '5c).Find the ending price for the current ticker.
               If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
                   endingPrice = Cells(j, 6).Value
               End If
            Next j

'6. Output the data for the current ticker.
        Sheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

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

'Check if the buttons are working
Sub ClearWorksheet()
    Worksheets("All Stocks Analysis").Activate
    Cells.Clear
End Sub