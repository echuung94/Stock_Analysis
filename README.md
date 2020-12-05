# stock-analysis

## Overview of Project - explain purpose of the analysis 

Steve is wanting to do additional research for his parents and wants to expand the dataset to include the stock 
market over the last few years. The purpose of the analysis was to refactor the original code to loop through all the 
data one time in order to collect the same information in a faster amount of time. In order to accomplish this, Visual Basic 
Application (VBA) in Excel was used to find the total daily volume and return. 

## Results 

### Refactoring the Code 
To create a more efficient output of my code, I used a nested for loop to get even more out of a for loop. Before creating the for loop, four different arrays were first 
created - tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices. The tickers array indicated the ticker symbol of a stock shown in the 2017 and 2018
worksheets. A tickerIndex variable was also created to access the correct index across the four different arrays. Another for loop was then created to loop over all the rows 
in the spreadsheet. Using the tickerIndex variable as the index allowed me to assign each of the different arrays to each ticker symbol listed in the worksheets 
before iterating through the dataset. 
   
   ```
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
    For i = 0 To 11
        tickerVolumes(i) = 0
        
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    
    For j = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
            If Cells(j - 1, 1).Value <> tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
                
            End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
            If Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(j, 6).Value

            '3d Increase the tickerIndex.
            
            tickerIndex = tickerIndex + 1
```

### Run Time for Each Year 
#### Original 
![Original_code_2017](https://github.com/echuung94/stock-analysis/blob/main/Resources/Original%20code%202017.png)
![Original_code_2018](https://github.com/echuung94/stock-analysis/blob/main/Resources/Original%20code%202018.png)

#### Refactor 
![Refactor_2017](https://github.com/echuung94/stock-analysis/blob/main/Resources/Refactor%202017.png)
![Refactor_2017](https://github.com/echuung94/stock-analysis/blob/main/Resources/Refactor%202018.png)

Based on the output run times, the refactored code ran approximately 0.5 seconds faster than the original code, making it more efficient than the original. 

![all_stocks__2017](https://github.com/echuung94/stock-analysis/blob/main/Resources/all%20stocks%202017%20.png)
![all_stocks__2018](https://github.com/echuung94/stock-analysis/blob/main/Resources/all%20stocks%202018.png)


## Summary: In a summary statement, address the following questions.

1. What are the advantages or disadvantages of refactoring code?
2. How do these pros and cons apply to refactoring the original VBA script?
