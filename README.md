# All Stocks Analysis Refactored

## Purpose 

Purpose of the study was to determine how to improve efficiencies of VBA coding. The orginal code that was completed for Steve to determine daily volume and return was then edited, or refractored, to improve the time to assess the data. With this new coding approach Steve will have the ability to better assess return on stocks for himself and his parents going forward.

## Results


**2017 Stock Tickers**

![alt text](https://github.com/CCoelho372/stock-analysis/blob/main/Challenge%202%20Ticker%20Chart_2017.png)

**2018 Stock Tickers**

![alt text](https://github.com/CCoelho372/stock-analysis/blob/main/Challenge%202%20Ticker%20Chart_2018.png)

 Based off of the code below the return was not able to be determined. By looking at the comparison charts on the challenge page, ENPH and RUN had positive returns in both 2017 and 2018. DQ had the largest return in 2017 and the the largest loss in 2018. It also was near the bottom in total daily volume. Thus showing that the ticker price may not be the most accurate resulting in that high variability.

  *Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1*


**2017 and 2018 Timer Results**

  *Original Script*

  ![alt text](https://github.com/CCoelho372/stock-analysis/blob/main/Assignment%202%20Ticker%20Chart_2017.png)

  ![alt text](https://github.com/CCoelho372/stock-analysis/blob/main/Assignment%202%20Ticker%20Chart_2018.png)

  *Refactored Script*

  ![alt text](https://github.com/CCoelho372/stock-analysis/blob/main/Challenge%202%20Ticker%20Timer_2017.png)

  ![alt text](https://github.com/CCoelho372/stock-analysis/blob/main/Challenge%202%20Ticker%20Timer_2018.png)

In comparison between the assignment and challenge script there were some great improvements in the length of time to complete the macro. In both the 2017 and 2018 script timers there was around 1/3 of a second taken off the time to run the script. As thousands of more tickers are assessed that amount of time saved will greatly increase the efficiency to conduct these analysis. 

##Summary

**Advantages and Disadvantages of Refactoring Code

The advantage of refactoring code is that you can save time and improve efficiencies in code. As data sets grow bigger, that time saving will continue to grow and have more of an effect than just a third of a second. The disadvantages of refactoring code is that you better know what you are doing. Case and point I took code that worked refactored it and then couldn't figure out how to show the returns. Sometimes if it isn't broken don't fix it.

**How is it Applied to Original VBA Script

In this case refactoring will be helpful for Steve and his parents as the datasets get larger. Currently there is not much difference between .4 seconds and .1 second to run a scipt after being refactored. but when you start analyzing 10x the amount of data that will start making a bigger difference.


## Refactored Code with Challenge Starter Download


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
    Dim tickerIndex As String
        tickerIndex = 0
    

    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    
    Dim tickerStartingPrices(12) As Single
    
    Dim tickerEndingPrices(21) As Double
    
       
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
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
            
          If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
          'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
           If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            '3d Increase the tickerIndex.
                
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then tickerIndex = tickerIndex + 1
        
           'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1
                
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
