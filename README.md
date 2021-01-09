# Stock Analysis
## Overview of Project
Steve is a financial advisor, and his parents want his help to invest in green energy they are particularly interested in DAQO New Energy (DQ). Steve came to me to assist him in analyzing 12 green energy stocks including DQ. I used VBA to help him perform this analysis and we determined that DQ is not the best option for Steve’s parents. He would now like me to modify this code so that if he wanted to analyze the entire stock market over the past few years he could do so in a timely manner. 

### Purpose
Create a code that can be executed in a timely manner so that Steve can analyze the entire stock market. 

## Results
Comparing the 12 green energy stocks that Steve wants to look at, it is evident that in 2017 all stocks performed well except for TERP however in 2018 the only stocks that performed well were ENPH and RUN. Based on these results the stock that Steve's parents are interested in, DQ, performed well in 2017 as the price increased 199.4%. However, in 2018 DQ did not perfrom well as the stock price dropped 62.6%. 

The original code that I created for Steve to analyze the 12 energy stocks took 0.773 seconds and 0.843 seconds to execute for the years 2017 and 2018, respectively.



'''

Worksheets("All Stocks Analysis").Activate

Dim startTime As Single
Dim endTime As Single

YearValue = InputBox("What year would you like to run the analysis on?")

startTime = Timer

  'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + YearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    'Assign Arrays and Indexes
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
    
    'initalize variables
    Dim startingPrice As Double
    Dim endingPrice As Double
    
    Worksheets(YearValue).Activate

    'find the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    'loop through tickers
    For i = 0 To 11
    
        ticker = tickers(i)
        totalVolume = 0
        
        'loop through rows
        Worksheets(YearValue).Activate
        For j = 2 To RowCount
        
            'total volume
            If Cells(j, 1).Value = ticker Then
            
                totalVolume = totalVolume + Cells(j, 8).Value
                
                End If
                
            'starting price
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            
                startingPrice = Cells(j, 6).Value
                
                End If
            
            'ending price
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            
                endingPrice = Cells(j, 6).Value
                
                End If
                
        Next j
        
          'Output Data
    
            Worksheets("All Stocks Analysis").Activate
    
            Cells(4 + i, 1).Value = ticker
            Cells(4 + i, 2).Value = totalVolume
            Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
        
    Next i
    
      'Formatting
         Range("A3:C3").Font.Bold = True
         Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
         Range("B4:B15").NumberFormat = "#,##0"
         Range("C4:C15").NumberFormat = "0.0%"
         Columns("B").AutoFit

        dataRowStart = 4
        dataRowEnd = 15

        For p = dataRowStart To dataRowEnd
    
        If Cells(p, 3).Value > 0 Then
    
            Cells(p, 3).Interior.Color = vbGreen
        
        
        ElseIf Cells(p, 3).Value < 0 Then
        
            Cells(p, 3).Interior.Color = vbRed
        
        
        Else: Cells(p, 3).Interior.Color = xlNone
        
        End If
        
     Next p
    
    endTime = Timer
    
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (YearValue)
    
'''

'''
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    YearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + YearValue + ")"
    
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
    Worksheets(YearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    Dim tickerIndex As Integer
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
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
       
        'End If
        End If
          
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
        
            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        'End If
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
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (YearValue)

End Sub

'''


