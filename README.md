# stock-analysis

## Overview of Project
This analysis aims at analyzing the entire dataset of each year's stock performance in a more efficient way by using refactoring in VBA. This code will loop through all the data one time in order to collect the stock ticker, volume and return information. Refactoring of the code makes the VBA script run faster, thus, can be applied to more stocks with more data by taking fewer steps of coding process, using less memory, or improving the logic of the code.




## Results
### Stock Performance
The images below are stock performance for the year of 2017 and 2018.

<img width="250" alt="Stock_Performance_2017" src="https://user-images.githubusercontent.com/82549066/116474565-ad15db80-a846-11eb-9adb-dfae9293c626.png">

<img width="248" alt="Stock_Performance_2018" src="https://user-images.githubusercontent.com/82549066/116474588-b2732600-a846-11eb-9ce3-f88b03062bbc.png">

From those tables, we can found that the overall stock performance in 2017 was much better than that in 2018 with 11 tickers had positive return in 2017 and only 2 tickers had positive return in 2018. The ticker that had best performance in 2017 was DQ that had a positive return of 199.4% and the one had best performance in 2018 was RUN that had 84% of return. The ticker that had the largest total daily volume in 2017 and 2018 are respectively SPWR and ENPH. In conclusion, the year of 2017 had better stock performance.


### Running Time
The images below illustrate the running time for our original code for both 2017 and 2018.
<img width="587" alt="VBA_Challenge_2017_Original" src="https://user-images.githubusercontent.com/82549066/116475595-10ecd400-a848-11eb-9a27-f8de4020d19c.png">

<img width="587" alt="VBA_Challenge_2018_Original" src="https://user-images.githubusercontent.com/82549066/116475603-13e7c480-a848-11eb-857d-b1aa73d39494.png">

The images below illustrate the running time for our refactoring code for both 2017 and 2018.
<img width="587" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/82549066/116475652-2661fe00-a848-11eb-8dfe-ae24ef44ea88.png">

<img width="587" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/82549066/116475657-295cee80-a848-11eb-9ae9-076c3883361e.png">

From those images, we can see that the running time becomes shorter for our refactoring codes by around 0.20312505 seconds for both 2017 and 2018. A variable called tickerIndex was created in our refactoring code for 4 different arrays-tickers, tickerVolumes, tickerStartingPrices and tickerEndingPrices. This allows to assign tickers, tickerVolumes, tickerSatrtingPrices and tickerEndingPrices to each stock before iterating and saves the running time. Detaied original code and refactoring code can be found below.

Original Code

    Sub yearValueAnalysis()
    Dim startTime As Single
    Dim endTime  As Single
    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer

    '1) Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stock Analysis").Activate

    Range("A1").Value = "All Stocks (" + yearValue + ")"

    'Add Headers
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    '2) Initialize array of all tickers
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
        totalVolume = 0
    '5) loop through rows in the data
    Worksheets(yearValue).Activate
    For j = 2 To RowCount
        '5a) Get total volume for current ticker
        If Cells(j, 1).Value = ticker Then
            totalVolume = totalVolume + Cells(j, 8).Value
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
        '6) Output data for current ticker
        Worksheets("All Stock Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

    Next i

    endTime = Timer
        MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    
    End Sub


Refacting Code

    Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stock Analysis").Activate
    
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
    For i = 0 To 11
    tickerVolumes(i) = 0
    
    Next i
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
        
        '3a) Increase volume for current ticker
       tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
         If Cells(i - 1, 1) <> tickers(tickerIndex) And Cells(i, 1) = tickers(tickerIndex) Then
         tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            
        'End If
        End If
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
         If Cells(i + 1, 1) <> tickers(tickerIndex) And Cells(i, 1) = tickers(tickerIndex) Then
         tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            

            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        'End If
        End If
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stock Analysis").Activate
        tickerIndex = i
        Cells(4 + i, 1).Value = tickers(tickerIndex)
        Cells(4 + i, 2).Value = tickerVolumes(tickerIndex)
        Cells(4 + i, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1
        
    Next i
    
    'Formatting
    Worksheets("All Stock Analysis").Activate
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




## Summary
1.What are the advantages or disadvantages of refactoring code?

Advantages
- Flexibility
It can improve the flexibility of the code which enables you to incooperate more functions.
- Maintability
The code becomes easier to read and maintain, saving the running time.

Disadvantages
- Run out of Time
You may spend a lot of time refactoring the code.
- Mistakes
You may create some mistakes or make the code more complex for refactoring.

2.How do these pros and cons apply to refactoring the original VBA script?

Pros
Our refactoring code becomes easier to read, saving the running time and more flexible.

Cons
It takes us the whole afternoon for refactoring which is time consuming.
