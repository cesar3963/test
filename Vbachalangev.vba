Sub sheetloopr()



Dim ws As Worksheet


For Each ws In Worksheets
    ws.Activate
    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 10) = "Yearly Charge"
    ws.Cells(1, 11) = "Precent Change"
    ws.Cells(1, 12) = "Total Stock Volume"

        ws.Range("a2", Range("A2").End(xlDown)).Copy
        ws.Range("I2").PasteSpecial
         Application.CutCopyMode = False
         
            ws.Range("G2", Range("G2").End(xlDown)).Copy
            ws.Range("l2").PasteSpecial
            Application.CutCopyMode = False
                
                Dim year As Double
                year = 2
                Do While ws.Cells(year, 1) <> ""
                ws.Cells(year, 10) = Cells(year, 3) - Cells(year, 6)
                ws.Cells(year, 11) = Cells(year, 10).Value
                ws.Cells(year, 11).NumberFormat = "0.0%"
                
                
                If ws.Cells(year, 10).Value >= 0 Then
                 ws.Cells(year, 10).Interior.ColorIndex = 4
                 
                 ElseIf ws.Cells(year, 10).Value < 0 Then
                 ws.Cells(year, 10).Interior.ColorIndex = 3
                 
                 End If
                 
                 year = year + 1
                Loop
                Columns("A:A").EntireColumn.AutoFit
                
                
                 
        
        

        
Next



End Sub


