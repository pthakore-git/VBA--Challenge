Sub stocks()
    Dim ticker As String
    
    Dim Ticker_Total As Double
    Ticker_Total = 0
    Dim openPrice As Double
    Dim closePrice As Double
    Dim change As Double
    Dim Summary_Table_index As Integer
    Summary_Table_index = 2
    Dim ws As Worksheet
    
    For Each ws In Worksheets
    

        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'MsgBox (Summary_Table_index)
        ws.Range("I1").EntireColumn.Insert
        ws.Range("I1") = "Ticker"
        ws.Range("J1").EntireColumn.Insert
        ws.Range("J1") = "Total Change"
        ws.Range("K1").EntireColumn.Insert
        ws.Range("K1") = "Percent Change"
        ws.Range("L1").EntireColumn.Insert
        ws.Range("L1") = "Total Stock Volume"
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("O4") = "Greatest Total Volume"
        
        For i = 2 To LastRow
        openPrice = ws.Cells(i, 3).Value
        'MsgBox ("Inside For loop" & " " & i)
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
               ' MsgBox ("Inside If" & " " & i)
                closePrice = ws.Cells(i, 6).Value
                ticker = ws.Cells(i, 1).Value
                Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
                'MsgBox ("Inside If " & Ticker_Total)
                ws.Range("I" & Summary_Table_index).Value = ticker
                ws.Range("L" & Summary_Table_index).Value = Ticker_Total
                ws.Range("J" & Summary_Table_index).Value = closePrice - openPrice
                'change = ws.Range("J" & Summary_Table_index).Value
                If ws.Range("J" & Summary_Table_index).Value > 0 Then
                    ws.Range("J" & Summary_Table_index).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & Summary_Table_index).Interior.ColorIndex = 3
                End If
                                
                ws.Range("K" & Summary_Table_index).Value = (closePrice - openPrice) / openPrice
                ws.Range("K" & Summary_Table_index).NumberFormat = "0.00%"
                Summary_Table_index = Summary_Table_index + 1
               ' MsgBox ("SumTableIndex: " & Summary_Table_index)
                Ticker_Total = 0
                
                
        
            Else
            Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
            'MsgBox ("Inside Else" & Ticker_Total)
            
            'MsgBox ("SumTableIndex: " & Summary_Table_index)
            End If
            
        Next i
        
        
        'MsgBox (ws.Name)
        'MsgBox (LastRow)
        ws.Range("Q2") = WorksheetFunction.Max(ws.Range("K2:K" & Summary_Table_index))
        ws.Range("Q3") = WorksheetFunction.Min(ws.Range("K2:K" & Summary_Table_index))
        ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & Summary_Table_index))
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
        increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & Summary_Table_index)), ws.Range("K2:K" & Summary_Table_index), 0)
        decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & Summary_Table_index)), ws.Range("K2:K" & Summary_Table_index), 0)
        volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & Summary_Table_index)), ws.Range("L2:L" & Summary_Table_index), 0)
        
        
        ws.Range("P2") = ws.Cells(increase_number + 1, 9)
        ws.Range("P3") = ws.Cells(decrease_number + 1, 9)
        ws.Range("P4") = ws.Cells(volume_number + 1, 9)
       
        
        
        
        Summary_Table_index = 2
        ws.Columns("I:O").AutoFit
    Next ws

End Sub
