# RBC-Module-2-Stock-Analysis

## Overview of Project
The purpose of this project was to help steve perform stock analysis for all companies, especially DAQO New Energy Corporation which makes silicion wafers for solar panels. Since his parents have invested all there money into DAQO.

- [Link to VBA_Challenge.xlsm](VBA_Challenge.xlsx)
- [Link to VBA_Challenge.vbs](VBA_Challenge.vbs)

### Purpose
The purpose of this project was to perform analysis on stocks for our friend steve while simulataneously working on a refactored (optimized and better structured) version of code for performing the analysis. Here we loop through all the rows of the dataset once without having to go through it multiple times for each ticker. This refractor speeds up the code significantly since we only need to go through the each row once.

## Comparison and Results of Refactored Code

### Comparison
To understand our results we must first understand the main difference in our code. Here are two code blocks providing an understanding as to how the code worked for the Initial Stock Analysis, using the `AllStocksAnalysis()` sub in Module1 of our VBA code and Refractored Stock Analysis, using the `AllStocksAnalysisRefactored()` sub in Module2 of our VBA code.

- Code for `AllStocksAnalysis()`
```
    '3a) Initializing variables
    Dim startingPrice As Single
    Dim endingPrice As Single
    '3b) Creating an array with variable tickers
    Dim tickers(11) As String
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
    
    '4) Get the number of rows to loop over
    RowCount = ws(Rows.Count, "A").End(xlUp).Row
    
    '5) looping through tickers
    For i = 0 To 11
             
        ticker = tickers(i)
        totalVolume = 0
                
            '6) looping through rows
            For j = 2 To RowCount
                    
                '6a) Getting the total volume for current ticker
                If ws(j, 1).Value = ticker Then
                        
                    totalVolume = totalVolume + ws(j, 8).Value
                
                End If
                    
                '6b) Getting the starting price for the current ticker
                If ws(j - 1, 1).Value <> ticker And ws(j, 1).Value = ticker Then
                    
                    startingPrice = ws(j, 6).Value
                        
                End If
                    
                '6c) Getting the ending price for the current ticker
                If ws(j + 1, 1).Value <> ticker And ws(j, 1).Value = ticker Then
                    
                    endingPrice = ws(j, 6).Value
                        
                End If
                    
            Next j
                
        '7) Output data for current ticker
        ws4(i + 4, 1).Value = ticker
        ws4(i + 4, 2).Value = totalVolume
        ws4(i + 4, 3).Value = (endingPrice / startingPrice) - 1
            
    Next i
```

- Code for `AllStocksAnalysisRefactored()`
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

    'Get the number of rows to loop over
    RowCount = ws(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    Dim tickerIndex As Integer
    tickerIndex = 0
        
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
    
    ''2b) Loop over all the rows in the spreadsheet
    For j = 2 To RowCount
            
        '3a) Increase volume for current ticker
        If ws(j, 1).Value = tickers(tickerIndex) Then
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + ws(j, 8).Value
                
        End If
                
        '3b) Check if the current row is the first row with the selected tickerIndex
        If ws(j - 1, 1).Value <> tickers(tickerIndex) And ws(j, 1).Value = tickers(tickerIndex) Then
                
            tickerStartingPrices(tickerIndex) = ws(j, 6).Value
                
        End If
                
        '3c) check if the current row is the last row with the selected ticker
        If ws(j + 1, 1).Value <> tickers(tickerIndex) And ws(j, 1).Value = tickers(tickerIndex) Then
                
            tickerEndingPrices(tickerIndex) = ws(j, 6).Value
                    
        '3d) Increase the tickerIndex
            
            tickerIndex = tickerIndex + 1
    
        End If
                
    Next j
            
    '4) Code to output data for current ticker(i), "Total Daily Volume", and "Return"
    For i = 0 To 11
        ws4(4 + i, 1).Value = tickers(i)
        ws4(4 + i, 2).Value = tickerVolumes(i)
        ws4(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
```
### Results of Refactored Code
<table align="center">
  <tr>
    <th>Runtime of Original Code</th>
    <th>Runtime of Refactored Code</th>
  </tr>
  <tr>
    <td><img src="https://github.com/mubeenkh4u/RBC-Module-2-Stock-Analysis/blob/main/Resources/VBA_Module_2017.png"></td>
    <td><img src="https://github.com/mubeenkh4u/RBC-Module-2-Stock-Analysis/blob/main/Resources/VBA_Challenge_2017.png"></td>
  </tr>
  <tr>
    <td><img src="https://github.com/mubeenkh4u/RBC-Module-2-Stock-Analysis/blob/main/Resources/VBA_Module_2018.png"></td>
    <td><img src="https://github.com/mubeenkh4u/RBC-Module-2-Stock-Analysis/blob/main/Resources/VBA_Challenge_2018.png"></td>
  </tr>
</table>

## Analysis of Stocks
After running the VBA code to perform the Stock Analysis. we can see that the stocks generally did better in 2017 as compared to 2018, with the exceptions of **RUN** and **TERP**, which saw significant increases.
<table align="center">
  <tr>
    <th>All Stocks (2017)</th>
    <th>All Stocks (2018)</th>
  </tr>
  <tr>
    <td><img src="https://github.com/mubeenkh4u/RBC-Module-2-Stock-Analysis/blob/main/Resources/AllStocksAnalysis2017.png"></td>
    <td><img src="https://github.com/mubeenkh4u/RBC-Module-2-Stock-Analysis/blob/main/Resources/AllStocksAnalysis2018.png"></td>
  </tr>
</table>

## Summary
### Pros of refactoring code
- In conclusion, refactoring of code is also known as restructuring of code for optimization of output without changing the output itself. This essentially means that the output is unchanged while the code that is being run has been modified to be more efficient/faster and extensible.
- From the above screenshots we can safely see that there is a slight difference in the `AllStocksAnalysis()` and `AllStocksAnalysisRefactored()` subroutines, with an approximated difference of 0.8s. These gains will be significantly larger if our dataset had hundreds and thousands of tickers and corresponding data.
- Lessens the repeatation of code.

### Cons of refactoring code
- Increasing the efficiency of the code by refactoring can cause the developer to lose precious time and can sometimes be not worth the effort, especially in case of small datasets.
- Requires a sound mind to chart out a map for the re-structuring of code to enable extensibility.
- In rare situations, this can lead to a deadlock where the developer might not know where to go.
