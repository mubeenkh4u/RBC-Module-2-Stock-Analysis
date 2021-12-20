Attribute VB_Name = "Module2"
Sub AllStocksAnalysisRefactored()

    Dim startTime As Single
    Dim endTime As Single
    Set ws4 = Sheet4.Cells
    
    yearValue = InputBox("What year would you like to run the analysis on?")
    
    startTime = Timer

    'Clear worksheet at the start of subroutine
    Call Module1.ClearWorksheet

    'Format the output sheet on All Stocks Analysis worksheet
    ws4(1, 1).Value = "All Stocks (" + yearValue + ")"
    ws4(1, 1).Font.Bold = True

    'Activate Data Sheet
    If yearValue = 2017 Then
    Set ws = Sheet1.Cells
    End If

    If yearValue = 2018 Then
    Set ws = Sheet2.Cells
    End If
    
    If yearValue = "" Then
    MsgBox ("Please select an year for Analysis")
    Exit Sub
    End If
    
    'create a header row
    ws4(3, 1).Value = "Year"
    ws4(3, 2).Value = "Total Daily volume"
    ws4(3, 3).Value = "Return"

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
                    
        '3d) Increase the tickerIndex.
            
            tickerIndex = tickerIndex + 1
    
        End If
                
    Next j
            
    '4) Code to output data for current ticker(i), "Total Daily Volume", and "Return"
    For i = 0 To 11
        ws4(4 + i, 1).Value = tickers(i)
        ws4(4 + i, 2).Value = tickerVolumes(i)
        ws4(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
 
    'Formatting
    Set fs = Sheet4
    fs.Range("A3:C3").Font.FontStyle = "Bold"
    fs.Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    fs.Range("B4:B15").NumberFormat = "#,##0"
    fs.Range("C4:C15").NumberFormat = "0.0%"
    fs.Columns("B").AutoFit
        
    dataRowStart = 4
    dataRowEnd = 15
    
    For i = dataRowStart To dataRowEnd

        If ws4(i, 3) > 0 Then

            'Color the cell green
            ws4(i, 3).Interior.Color = vbGreen

        ElseIf ws4(i, 3) < 0 Then

            'Color the cell red
            ws4(i, 3).Interior.Color = vbRed

        Else

            'Clear the cell color
            ws4(i, 3).Interior.Color = xlNone

        End If

    Next i

endTime = Timer
MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub


