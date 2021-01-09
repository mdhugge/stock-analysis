# Stock Analysis
## Overview of Project
Steve is a financial advisor and his parents want his help to invest in green energy, they are particularly interested in DAQO New Energy (DQ). Steve came to me to assist him in analyzing 12 green energy stocks including DQ. I used VBA to help him perform this analysis and we determined that DQ is not the best option for Steve’s parents. He would now like me to modify this code so that if he wanted to analyze the entire stock market over the past few years he could do so in a timely manner. 

### Purpose
Refactor code so that it can be executed in a timely manner for Steve to analyze the entire stock market. 

## Results
Comparing the 12 green energy stocks that Steve wants to look at, it is evident that in 2017 all stocks performed well except for TERP however in 2018 the only stocks that performed well were ENPH and RUN. Based on these results the stock that Steve's parents are interested in, DQ, performed well in 2017 as the price increased 199.4%. However, in 2018 DQ did not perfrom well as the stock price dropped 62.6%. 

The original code that I created for Steve to analyze the 12 energy stocks took 0.773 seconds and 0.843 seconds to run for the years 2017 and 2018, respectively.

![VBA_Challenge_2017_Original](https://github.com/mdhugge/stock-analysis/blob/main/Resources/VBA_Challenge_2017_Original.png)

![VBA_Challenge_2018_Original](https://github.com/mdhugge/stock-analysis/blob/main/Resources/VBA_Challenge_2018_Original.png)

The original code included nested for loops which can increase the time required to run the code. The outer loop is executed 12 times and each time the outer loop executes the inner loop executes 3012 times. As a result there are 12 x 3012 executions in this code. 
    
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

The refactored code I created did not include nested loops and it was faster. It ran in 0.140 seconds and 0.137 seconds for 2017 and 2018, respectively. 

![VBA_Challenge_2017](https://github.com/mdhugge/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

![VBA_Challenge_2018](https://github.com/mdhugge/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

The refactored code used 4 arrays; tickers, tickerVolumes, tickerStartingPrices and tickerEndingPrices. The tickerIndex variable was used to access an index accross these 4 arrays. 

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

An advantage of refactoring a code is that the code becomes more efficient. The actual functionality of the code is not changed, but it can be executed in a fewer number of steps and that may result in faster execution. Another advantage is that refactoring may improve the logic of the code, consequently making it easier for someone to read and understand. When a code is first written it may not be the best way to accomplish the task and refactoring will clean up the code.  

A disadvantage of refactoring code is that it can be a time-consuming process to produce the same functionality. Moreover, if someone is being paid to write the code refactoring would be additional money spent.

- How do these pros and cons apply to refactoring the original VBA script?

Refactoring the VBA script increased the efficiency as the refactored code ran faster than the original. The refactored code does not have as many steps, since the nested for loop is replaced with 4 arrays and an index variable. I found that refactoring the code was challenging and there was a lot of trial and error to produce the same end result. The refactored code was 0.63 seconds and 0.70 seconds faster for the years 2017 and 2018, respectively. Personally, I do not think the amount of time I spent to refactor the code is appropriately reflected in this faster execution. However, I only used this code to analyze 12 stocks so perhaps when the same code is used to analyze a greater number of stocks the efficiency of the refactored code will be more applicable. As a result if Steve uses the refactored code to analyze the entire stock market it would run more quickly than if he used the original code. 




