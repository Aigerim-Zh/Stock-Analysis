# Green Energy Stock Analysis using Excel VBA scripting

Author: Aigerim Zhanibekova
Email: azhanibekova787@gwu.edu
Date: January 2022
Instituition: The George Washington University

# Overview of the Project

## Background 

Fossil fuels are limited and will eventually be used up. There is more and more reliance on alternative energy production. Green energy is gaining more and more popularity. At the same time, green energy is becoming more attractive to investors due to falling costs of renewable energy production. There are many forms of green energy to invest in, including hydroelctricty, windenergy, geothermal energy and bio energy (DiLallo, 2021). 

## Purpose

- The purpose of this project is to perform data analysis to determine which green energy stocks might be most beneficial to invest in. 

- After the analysis is run, the project further aims to refactor the existing VBA code to make it more efficient and readible. This step is important in the case the analysis is run on larger datasets. The main measure of efficiency improvement is the running time of the code. 

## Dataset

The current dataset consists of daily data for 12 green energy stocks 2017 and 2018 years. For each stock, there are 251 rows representing business days in each year. The variables include the ticker name of stocks, date, open and closing price, high and low price, adjusted low price and total daily volume.

# Analysis
There are two files in this repository with the same results. However, the code in each file is different in terms of its applicability and efficiency. 

* In the [**"Stocks_Analysis_Before_Refactoring.xlsm"** ](https://github.com/Aigerim-Zh/Stock-Analysis/blob/main/Stocks_Analysis_Before_Refactoring.xlsm) file, you can find the analysis run with the original code. 
* In the [**"Stocks_Analysis_After_Refactoring.xlsm"**](https://github.com/Aigerim-Zh/Stock-Analysis/blob/main/Stocks_Analysis_After_Refactoring.xlsm)) file, you can find the analysis with the refactored code. 
* In each file, there are two sheets with the stock data we are using for the analysis, named "2017" and "2018". 
* In each file, the results can be accessed in the "All Stocks Analysis" sheet through **macro-enable buttons** for analysis and formatting. 

## Total Daily Volume and Yearly Return for Each Stock

- Total Daily Volume is calculated as a sum of all daily volumes for each stock. This variable can give us a rough idea of how often the stock was traded in a particular year. 
    - Daily Volume (column H) in both "2017" and "2018" worksheets shows how actively a stock was traded, i.e., what number of shares was traded throughout the day. 
    - It is reasonable to assume that if a stock is traded often, the stock value might be reflected in its price.

- Yearly Return is calculated as as percentage change between the price at the beginning of the year and at the end of the year. If it is positive, this means it might be profitable to invest in this stock, also given that it was traded enough in that year. 
 
### Stock Analysis Results 
![2017](https://github.com/Aigerim-Zh/Stock-Analysis/blob/main/Green_Stocks_Dataset.xlsx)
![2018](https://github.com/Aigerim-Zh/Stock-Analysis/blob/main/Resources/2018AnalysisResults.png)

In 2017, all stocks generated positive returns except for the TERP stock. 
In 2018, all stocks generated negative returns except for the ENPH and RUN stocks. These two stocks showed positive returns for both years and were traded very frequently as well. 

### VBA Code for the Analysis Explained

#### _Recording Code Running Time_
To assess the efficiency of a code, we need to know how long it takes to execute it. To do that, starting and ending time variables are created in each code. 

```
Sub AllStocksAnalysis()
    Dim startTime As Single
    Dim endTime As Single
End Sub
```
#### _Customizing Analysis Year_

We have data for 2017 and 2018 years in separate sheets. Instead of having separate macros for each year, we can let the user decide for which year to access the results. 

```
    yearValue = InputBox("What year would you like to run the analysis on?")
```

#### _Setting a Timer_

To set a timer, we set the startTime and endTime variables equal the timer function. It is important to set the startTime after the user chooses the analysis year and set the endTime at the very end of the code. 

```
    startTime = Timer
```
```
    endTime = Timer
```
#### _Preparing the "All Stocks Analysis" Output Sheet_

In the "All Stocks Analysis" sheet, we are creating a table with three columns - Ticker, Total Daily Volume, and Yearly Return. The code looks as follows:

```
    Sheets("All Stocks Analysis").Activate
        Range("A1").Value = "All Stocks (" + yearValue + ")"
        Cells(3, 1).Value = "Ticker"
        Cells(3, 2).Value = "Total Daily Volume"
        Cells(3, 3).Value = "Yearly Return"
```
#### _Ticker_

In the "All Stocks Analysis", the Ticker column is for the ticker abbreviation of the stock. There are 12 tickers in total. 

**The code before refactoring** does not have a way to produce a column with the unique ticker names, and it had to be done with a pivot table manually. **After refactoring**, a column with the uniqiue ticker names is exported in a separate sheet called "Unique Tickers"; this allows to copy the ticker names to save time on enetering them manually as shown below:

```
Sheets(yearValue).Activate
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row

    Range("A1:A" & rowEnd).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Sheets("Unique Tickers").Range("A1"), Unique:=True

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
```

#### _Calculating Total Daily Volume and Yearly Return_

**_Before Refactoring_ the code was organized as follows**:
- Total Daily Volume and Yearly Return are calculated using nested for loops. 
- Before we loop through rows in the data, we need to initialize the loop for identifying tickers. 
- The reason that totalVolume variable is set to zero inside the outer ticker loop is that it needs to be reset to zero after the loop has gone through one ticker and is moving to the next. 
- Before the inner loop starts, we need to make sure that the correct worksheet is activated with _Sheets(yearValue).Activate_

```
4. Loop through the tickers.

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

            #### _Outputing the results in the "All Stocks Analysis" sheet_

' Now, let's output the results in the "All Stocks Analysis" sheet. The Cells() function is used, so the output for each ticker is displayed on a new row. 

'6. Output the data for the current ticker.
        Sheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

    Next i
    'we need to make sure that we set the endTime variable in the end of the code. 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    
End Sub
```

**After refactoring, the code looks as follows**:
- Tickers, Total Daily Volume, and Yearly Return are defined through arrays and a tickerIndex to access the correct index across these arrays. 
- Setting the most correct type of arrays can also increase the efficiency of the code by using less memory. 

```
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
```
#### _Formatting the Table_
To improve reading of the table, some formatting is present in both VBA scripts.  

1. The headers of the table are made bold and have a bottom thin border. 

```
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
```
2. Total Daily Volume column is auto-fitted for its length and the values are formatted with a 1000 comma separator. Yearly Return values are formatted to percentages wih two digits of precision. 

```
    'Formatting the numbers
    Range("B4:B15").NumberFormat = "#,##"
    Range("C4:C15").NumberFormat = "0.00%"
    'Autofitting the total volume column
    Columns("B").AutoFit
```
3. Conditional formatting is applied to the Yearly Return column. The cells with positive returns are shaded green and those with negative with red. And, if the return is equal to zero, the cell has no fill. 

```
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
```

#### _Creating Buttons to Access Results without the Developer Tab_

Finally, in the "All Stocks Analysis" worksheet, there are two buttons created for the end-user to run the analysis without having to run the code from the Developer tab. 

- The "Run All Stock Analysis" button will run all the analysis performed [here](). 
- The "Add Formatting" button will add the formatting we added [here](). 
- The "Clear the Worksheet" button will clean the worksheet. To clean the worksheet to analyze another year, the following subroutine can be used:

```
'Check if the buttons are working
Sub ClearWorksheet()
    Worksheets("All Stocks Analysis").Activate
    Cells.Clear
End Sub

```
## Results for the VBA Code Performance Before and After Refactoring

As shown above, there are two VBA codes that were prepared for this analysis. [The first code](https://github.com/Aigerim-Zh/Stock-Analysis/blob/main/VBA_Code_Before_Refactoring.vbs) is applicable to the parameters of the current dataset. The code efficiency is X seconds of running time for the 2017 data, and Y seconds for the 2018 data. 
![](https://github.com/Aigerim-Zh/Stock-Analysis/blob/main/Resources/RunTime_2017_Before_Refactoring.png)
![](https://github.com/Aigerim-Zh/Stock-Analysis/blob/main/Resources/RunTime_2018_Before_Refactoring.png)

[The second code](https://github.com/Aigerim-Zh/Stock-Analysis/blob/main/VBA_Code_After_Refactoring.vbs) is a refactored code that can be applied to a larger dataset with even thousands more of stocks. In addition, it has greater efficiency. 
![](https://github.com/Aigerim-Zh/Stock-Analysis/blob/main/Resources/RunTime_2017_After_Refactoring.png)
![](https://github.com/Aigerim-Zh/Stock-Analysis/blob/main/Resources/RunTime_2018_After_Refactoring.png)

## Summary 
### Advantages of Refactoring Code
- Refactoring can make the code more efficient (reduce its running time) and also increase its applicapibilty to larger datasets. 
- Refactoring makes the code cleaner and more readible, which makes errors detection easier. 
### Disadvantages of Refactoring Code
- Given the fact that we already a running code and are trying to refactor it might open room for logical errors, i.e., it is important that it is refactored correctly. That is why it is a good practice to save the code before and after refactoring, as done in this project. 

### How Pros and Cons of Refactoring apply to the original VBA script in this project?
- Refactoring our code helped us to increase its efficiency. 
   - As shown above, the time decrease by 2.25% in 2017 and 3.6% in 2018. However, note that the running times change every time you run it again but, in general, the running times will be lower with the refactored code. 
   - Refactoring code also helped us to use less memory by setting arrays and their most correct types. 
- Refactoring made the code applicable to larger datasets. Even if the previous code might have run for larger datasets, it may take much more time to execute. 

# References 
DiLallo, M. (2021). Investing in Renewable Energy Stocks. Retrieved from https://www.fool.com/investing/stock-market/market-sectors/energy/renewable-energy-stocks/