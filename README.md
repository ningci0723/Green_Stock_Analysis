# Green Stock Analysis

## Overview of Project

Steve wants to conduct analysis on green stocks in 2017 and 2018 for his parents to see if it is worth investing in. In order to help him to conduct this analysis, I will use the Visual Basic Application in Excel to find the stock's total daily volume and annual return, which will help Steve to know which stock will be the best option for his parents.

### Purpose

The purpose of this analysis is editing, or refactoring the Module 2 solution code to find a more efficient way to look conduct analysis and find the best stock by using VBA. This project is refactoring the code to make the VBA script run faster.


## Results

### Refactoring the Code
In order to make the code more efficient, I created three different arrays which were tickerVolumes(12),tickerStartingPrices(12) and tickerEndingPrices(12). The tickers array created in the original code was used to establish the ticker symbol of a stock. I matched three different arrays which were tickerVolumes(12),tickerStartingPrices(12) and tickerEndingPrices(12) with the tickers array by using a variable named as tickerIndex. Here is the refactor code.

#### Refactored Code

```
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

    '1a) Create a ticker index to reference proper ticker in the arrays.
    Dim tickerIndex As Integer
    'Initiate tickerIndex at zero.
    tickerIndex = 0


    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single



    '2a) Create for loop to analyze each ticker in the array.
    For tickerIndex = 0 To 11
    'Initiate each ticker's volume at zero.
    tickerVolumes(tickerIndex) = 0

    'Activate data worksheet
    Worksheets(yearValue).Activate

        '2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount

            '3a) Increase volume for current ticker.
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value


            '3b) Check if the current row is the first row with the current ticker.

            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                'if it is the first row for current ticker, set starting price.
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value

            'End If
            End If


        '3c) Check if the current row is the last row with the current ticker.

            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                'if it is the last row for current ticker, set ending price.
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

            'End if
            End If

        '3d) Check if the current row is the last row with the current ticker.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then

                'if it is, increase tickerIndex to move on to next ticker in array.
                tickerIndex = tickerIndex + 1

            'End If
            End If

        Next i

    Next tickerIndex

    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.

    For i = 0 To 11

        'Activate Output Worksheet
        Worksheets("All Stocks Analysis").Activate

        'Ticker Row Label
        Cells(4 + i, 1).Value = tickers(i)

        'Sum of Volume
        Cells(4 + i, 2).Value = tickerVolumes(i)

        'ReturnValue
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1



    Next i

```

  Below is the Original Code
#### Original Code
  ```
  Initialize array of all tickers

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

'3a) Initialize variables for starting price and ending price

    Dim startingPrice As Double
    Dim endingPrice As Double

'3b) Activate data worksheet

    Worksheets(yearValue).Activate

'3c) Get the number of rows to loop over

    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

'4) Loop through tickers

    For i = 0 To 11
    ticker = tickers(i)
    TotalVolume = 0
    Worksheets(yearValue).Activate

'5) loop through rows in the data

For j = 2 To RowCount

    '5a) Find total volume for current ticker

    If Cells(j, 1).Value = ticker Then

        'increase totalVolume by the value in the current row
        TotalVolume = TotalVolume + Cells(j, 8).Value

End If

        '5b) Find starting price for current ticker

    If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        'set starting price
        startingPrice = Cells(j, 6).Value

    End If

        '5c) Find ending price for current ticker

        If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        'set ending price
        endingPrice = Cells(j, 6).Value

    End If

    Next j
'6) Output data for current ticker

    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = TotalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

 Next i
  ```

Below are the run time for Original code for 2017 and 2018
![Original_2017](https://github.com/ningci0723/Green_Stock_Analysis/blob/main/Original_2017.png)
![Orinigal_2018](https://github.com/ningci0723/Green_Stock_Analysis/blob/main/Original_2018.png)

Below are the run time for refactored code for 2017 and 2018
![Refactored_2017](https://github.com/ningci0723/Green_Stock_Analysis/blob/main/VBA_Challenge_2017.png)
![Refactored_2018](https://github.com/ningci0723/Green_Stock_Analysis/blob/main/VBA_Challenge_2018.png)
