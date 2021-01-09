# Stock Analysis
## Overview of Project
Steve is a financial advisor and his parents want his help to invest in green energy, they are particularly interested in DAQO New Energy (DQ). Steve came to me to assist him in analyzing 12 green energy stocks including DQ. I used VBA to help him perform this analysis and we determined that DQ is not the best option for Steve’s parents. He would now like me to modify this code so that if he wanted to analyze the entire stock market over the past few years he could do so in a timely manner. 

### Purpose
Refactor code so that it can be executed in a timely manner for Steve to analyze the entire stock market. 

## Results
Comparing the 12 green energy stocks that Steve wants to look at, it is evident that in 2017 all stocks performed well except for TERP however in 2018 the only stocks that performed well were ENPH and RUN. Based on these results the stock that Steve's parents are interested in, DQ, performed well in 2017 as the price increased 199.4%. However, in 2018 DQ did not perfrom well as the stock price dropped 62.6%. 

The original code that I created for Steve to analyze the 12 energy stocks took 0.773 seconds and 0.843 seconds to execute for the years 2017 and 2018, respectively.

![VBA_Challenge_2017_Original](https://github.com/mdhugge/stock-analysis/blob/main/Resources/VBA_Challenge_2017_Original.png)

![VBA_Challenge_2018_Original](https://github.com/mdhugge/stock-analysis/blob/main/Resources/VBA_Challenge_2018_Original.png)

The original code included nested loops which can increase the time required to execute the code. The outer loop is executed 12 times and each time the outer loop executes the inner loop executes 3012 times. As a result there are 12 x 3012 executions in this code. 
    
    'Loop through tickers
    For i = 0 To 11
    
        ticker = tickers(i)
        totalVolume = 0
        
        'Loop through rows
        Worksheets(YearValue).Activate
        For j = 2 To RowCount
        
            'Total volume
            If Cells(j, 1).Value = ticker Then
            
                totalVolume = totalVolume + Cells(j, 8).Value
                
                End If
                
            'Starting price
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            
                startingPrice = Cells(j, 6).Value
                
                End If
            
            'Ending price
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            
                endingPrice = Cells(j, 6).Value
                
                End If
          Next j      
    Next i

The refractored code I created did not include nested loops and it was faster. It was executed in 0.140 seconds and 0.137 seconds for 2017 and 2018, respectively. 

![VBA_Challenge_2017](https://github.com/mdhugge/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

![VBA_Challenge_2018](https://github.com/mdhugge/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

    'Create a ticker Index
    Dim tickerIndex As Integer
    tickerIndex = 0

    'Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    'Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
    
    Next i
    
    'Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        'Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

        'Check if the current row is the first row with the selected tickerIndex
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
       
        End If
          
        'Check if the current row is the last row with the selected ticker
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
            'Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        End If
        
    Next i

## Summary
- What are the advantages or disadvantages of refactoring code?

An advantage of refactoring a code is that the code becomes more efficient. The actual functionality of the code is not changed, but it can be executed in a fewer number of steps and that may result in faster execution. Another advantage is that refactoring may improve the logic of the code, consequently making it easier for someone to read and understand. When a code is first written it may not be the best way to accomplish the task. As a result, refactoring will clean up the code.  



- How do these pros and cons apply to refactoring the original VBA script?





