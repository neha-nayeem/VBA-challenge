Attribute VB_Name = "Module1"
Sub Stocks()

    '--- Define variables ---
    Dim Ticker As String
    Dim StockOpen, StockClose, YearlyChange, PercentChange As Double
    Dim i, LastRow, OutputRow As Long
    Dim Volume, TotalVolume As LongLong
    Dim GreatestIncrease, GreatestDecrease As Double
    Dim GreatestVolume As LongLong
    
    '--- Insert summary headers into worksheet ---
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    
    '--- Find the last row in the worksheet and assign to LastRow ---
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        'MsgBox (Str(LastRow))
    
    '--- Assign first ticker and opening price values from data---
    Ticker = Range("A2").Value
    StockOpen = Cells(2, 3).Value
    
    '--- Assign first output row for the summary table to be on row 2 ---
    OutputRow = 2
    
    '--- For loop to go through data from row 2 to the last row  ---
    For i = 2 To LastRow
                   
        '--- Conditional statement to check whether the next row contains the same ticker value i.e. the current ticker is continuing ---
        If Ticker = Cells(i + 1, 1).Value Then
            
            '--- Assign a value to volume from worksheet ---
            Volume = Cells(i, 7).Value
            
            '--- Calculate total stock volume ---
            TotalVolume = TotalVolume + Volume
            
        '--- Conditional statement to check whether the next row's ticker value is NOT the same i.e. the current ticker has ended ---
        ElseIf Ticker <> Cells(i + 1, 1).Value Then
        
            '--- Assign volume and closing price from this row (end of year value) from worksheet ---
            StockClose = Cells(i, 6).Value
            Volume = Cells(i, 7).Value
        
            '--- Calculate yearly change ---
            YearlyChange = StockClose - StockOpen
            
            '--- Calculate percent change ---
            PercentChange = (YearlyChange / StockOpen)
            
            '--- Calculate total stock volume ---
            TotalVolume = TotalVolume + Volume
            
            '--- Insert calculated values into summary table on worksheet ---
            Cells(OutputRow, 9).Value = Ticker
            Cells(OutputRow, 10).Value = YearlyChange
            Cells(OutputRow, 11).Value = PercentChange
            Cells(OutputRow, 12).Value = TotalVolume
            
            '--- Nested If statement to check YearlyChange and change cell colors (conditional formatting) ---
                
                '--- if YearlyChange is positive, then change the cell color to green ---
                If YearlyChange > 0 Then
                    Cells(OutputRow, 10).Interior.ColorIndex = 4
                
                '--- change color to red ---
                Else
                    Cells(OutputRow, 10).Interior.ColorIndex = 3
                    
                '--- end nested if statement
                End If
            
            'Assign greatest % increase, decrease and volumes to variables ---
            GreatestIncrease = Application.WorksheetFunction.Max(Range("K:K"))
            GreatestDecrease = Application.WorksheetFunction.Min(Range("K:K"))
            GreatestVolume = Application.WorksheetFunction.Max(Range("L:L"))
            
            '--- Nested If statement to check greatest % increase, decrease, volume and return value to worksheet ---
             If PercentChange = GreatestIncrease Then
                Range("P2").Value = Ticker
                Range("Q2").Value = GreatestIncrease
            
            ElseIf PercentChange = GreatestDecrease Then
                Range("P3").Value = Ticker
                Range("Q3").Value = GreatestDecrease
                
            ElseIf TotalVolume = GreatestVolume Then
                Range("P4").Value = Ticker
                Range("Q4").Value = GreatestVolume
                
            '--- end nested if statement for greatest values ---
            End If
           
            '--- Reassign values for the next ticker (new ticker value, new open price and reset total volume) ---
            Ticker = Cells(i + 1, 1).Value
            StockOpen = Cells(i + 1, 3).Value
            TotalVolume = 0
            
            '--- Update OutputRow so that info for new ticker is entered onto next row in the summary table ---
            OutputRow = OutputRow + 1
            
        '--- End conditional statements ---
        End If
        
    '--- Continue for loop to next row ---
    Next i
    
End Sub
