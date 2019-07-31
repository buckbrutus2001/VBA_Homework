

Sub Totalvolume()



For Each ws In Worksheets

    Dim RowCount As Long
    Dim Total As Double
    Dim TotalLineNumber As Long

    TotalLineNumber = 2
    RowCount = 0
    Total = 0
    RowCount = ws.Cells(Rows.Count, 1).End(xlUp).Row


        For i = 2 To RowCount

            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
                Total = Total + ws.Cells(i, 7).Value
                ws.Range("I" & TotalLineNumber).Value = ws.Cells(i, 1).Value
                ws.Range("J" & TotalLineNumber).Value = Total
                TotalLineNumber = TotalLineNumber + 1
                Total = 0
                
            Else

                Total = Total + ws.Cells(i, 7).Value

            End If
        
        Next i

        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Total Stock Volume"
        ws.Columns("A:J").AutoFit
        
Next ws
End Sub

