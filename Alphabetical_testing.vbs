Sub tickertape()
    ' Loop through all sheets
    Dim ws As Worksheet
    
    For Each ws In ActiveWorkbook.Worksheets
    SumTotal = 0
        
        lastrow = ws.Cells(ws.Rows.Count, "A").End(xlup).Row
        
        
        Dim Ticker_Name As String
        Dim Open_close As Double
            Open_close = 0
        Dim Summary_Table_Row As Double
        Summary_Table_Row = 2
        Dim open_value As Double
        open_value = ws.Cells(2, 3)
        
        Dim StockVolOpen As Range
        Set StockVolOpen = ws.Range("g2")
        
        
        Dim StockVolClose As Range
        
        
        Dim stockvol As Range
        
        
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Yearly Change"
        ws.Cells(1, 11) = "Percent Change"
        ws.Cells(1, 12) = "Total Stock volume"
        ws.Cells(2, 16) = "greatest % increase"
        ws.Cells(3, 16) = "greatest % Decrease"
        ws.Cells(4, 16) = "Greatest total volume"
        ws.Cells(1, 17) = "Ticker"
        ws.Cells(1, 18) = "Value"
        
        ws.Cells(2, 18) = WorksheetFunction.Max(Columns(11))
        ws.Cells(3, 18) = WorksheetFunction.Min(Columns(11))
        ws.Cells(4, 18) = WorksheetFunction.Max(Columns(12))
        
        
        
        
        
        'Set GetLastNo = ws.Range(ws.Cells(lLastRow, lLastCol).Address)
        'Range("stockvol").Value = Application.Sum(Range(Cells(13, 2), Cells(14, 2)))
        


        
        
        For i = 2 To lastrow
       If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
        
       Set StockVolClose = ws.Cells(i + 1, 7)
       SumTotal = SumTotal + StockVolClose
       
        
        ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker_Name = ws.Cells(i, 1).Value
            Open_close = ws.Cells(i, 6) - open_value
            ws.Range("L" & Summary_Table_Row).Value = SumTotal
            
            
            ws.Range("I" & Summary_Table_Row).Value = Ticker_Name

      ' Print the Brand Amount to the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = Open_close
      
      If open_value = 0 Or IsEmpty(open_value) Then
      Cells(i, 10).Value = "Null"
Else

    Dim percentchange As Double
    percentchange = (Open_close) / (open_value)
    ws.Range("K" & Summary_Table_Row).Value = FormatPercent(percentchange)

End If
  
        'Set stockvol = (ws.Range("StockVolOpen"))
        ' Total stock volume
        'StockVol = (ws.Cells(2, 7), ws.Cells(i, 7)))
        'Dim SumTotal As Double
        'SumTotal = Excel.WorksheetFunction.Sum(stockvol)
        SumTotal = 0
        
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Brand Total
      open_value = ws.Cells(i + 1, 3)
      
      
      
      
      End If
        
        Next i
    
Next ws
MsgBox ("all done!")
End Sub