# Stock Analysis in VBA

## Overview of Project

### Purpose of Analysis
The purpose of this analysis was to help our good friend Steve quickly review stock performance so he could better advice his parents. We made a code in VBA to quickly screen through stocks for 2017 and 2018 by looking at annual volume and return.

## Results

### Stock Performance between 2017 and 2018
First we wanted to look at the total daily volume and return for each stock so we made a code with a `For Loop` to run though each ticker index. We inserted a second `For Loop` into the first one, to check each rows from 2 to RowCount, which was determined using the formula: `RowCount = Cells(Rows.Count, "A").End(xlUp).Row.` The nested `For Loop` computed the total daily volume for the year, set a starting price and an ending price used to generate the annual return.
```
For i = 0 To 11
   ticker = tickers(i)
   totalVolume = 0
    'loop through rows in the data
   Worksheets(yearValue).Activate
   For j = 2 To RowCount
       'Get total volume for current ticker
        If Cells(j, 1).Value = ticker Then

        totalVolume = totalVolume + Cells(j, 8).Value

        End If
       'get starting price for current ticker
        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

        startingPrice = Cells(j, 6).Value

        End If
       'get ending price for current ticker
        If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

        endingPrice = Cells(j, 6).Value
        End If

   Next j
   
      'Output data for current ticker
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    Cells(4 + i, 4).Value = startingPrice
    Cells(4 + i, 5).Value = endingPrice
    
   
Next i
```

For this original code it is very important to activate the correct worksheet in each loop since the data to analyze and the output results are on different worksheet. We only had data for 2017 and 2018 but the code was made to be flexible in case other years are added. For this, we obtain the `yearValue` variable using an input box and referencing to it throughout the code. When we want to insert the `yearValue` text in a cell we used the `+ +` signs. 
```
Range("A1").Value = "All Stocks (" + yearValue + ")"
```
At the end of the code, we inserted a few lines to format the rows and columns of interest with bold text and various color highlights then stopped our timer.
```
Worksheets("All Stocks Analysis").Activate
Range("A1:E3").Font.Bold = True
Range("A3:E3").Borders(xlEdgeBottom).LineStyle = xlContinuous
Range("A3:E3").Font.Italic = True
Range("A1:E3").Interior.ColorIndex = 15
Range("B4:B15").NumberFormat = "#,##0"
Range("C4:C15").NumberFormat = "0.0%"
Range("D4:E15").NumberFormat = "$0.00"
Columns("A:E").AutoFit

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

    Next i
    
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    
End Sub
```

The original code was a little slow to run (~ 1 second) so we refactored it using a new `tickerIndex` variable. This allowed us to store data for each ticker in the memory without having to run through each row multiple times while avoiding nested loops. This new variable assigned an index across four arrays which are initialized as below:
```
tickerIndex = 0
Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single
```
Each array is then refered to in the code with the variable `(tickerIndex)`.
```
For i = 0 To 11

    tickerVolumes(i) = 0
    
Next i

    '2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
                
        '3b) Check if the current row is the first row with the selected tickerIndex.
        
        If Cells(i - 1, 1) <> Cells(i, 1) Then
            
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        
        If Cells(i + 1, 1) <> Cells(i, 1) Then
            
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
              
        '3d Increase the tickerIndex.
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
```
The new refactored code was almost ten times faster (~0.1 second) than the original code and may be very helpful for Steve as the dataset to analyze might just get bigger and bigger. 

#### Original Code 2017
![Code1_2017](Resources/Code1_2017.png)
#### Refactored Code 2017
![VBA_Challenge_2017](Resources/VBA_Challenge_2017.png)
#### Original Code 2018
![Code1_2018](Resources/Code1_2018.png)
#### Refactored Code 2018
![VBA_Challenge_2018](Resources/VBA_Challenge_2018.png)



Overall, 2017 was a great year to invest with all analyzed stocks in the green except for TERP with -7.2% return. On the other hand, 2018 could be considered a terrible year with 10 of the 12 stocks in the red with only ENPH and RUN having positive returns. While the dataset is small to make a clear conclusion, we would recommend Steve to invest in ENPH and RUN since these two stocks got positive returns even during a bad market year. This would suggest such companies are well managed, are keeping acceptable level of cashflow and are profitable even during uncertainties and economic downturn. 

#### Return 2017
![Return_2017](Resources/Return_2017.png)
#### Return 2018
![Return_2018](Resources/Return_2018.png)

## Summary

### Advantages and Disadvantages of Refactoring Code
The refactored code greatly increase efficiency and might be better for big dataset but doing so took extra labor and could affect cost. 

### Advantages and Disadvantages of the Original and Refactored VBA Script

In the original code we run through each row, generate data then output it right away before moving to the next ticker. In the refactored code, we run through each row one time while storing data in the memory for each ticker and output it all at same time at the end. The original code is slower but empties the memory after each ticker. We would need to test if the refactored code can handle thousands of tickers without crashing or running out of memory but for this purpose it works very well.
