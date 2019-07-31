

Sub Totalvolume()



For Each ws In Worksheets

    Dim RowCount As Long
    Dim Total As Double
    Dim TotalLineNumber As Long
    Dim OpeningPrice as Double
    Dim ClosingPrice as Double

    TotalLineNumber = 2
    RowCount = 0
    Total = 0
    OpeningPrice = 0
    ClosingPrice = 0
    RowCount = ws.Cells(Rows.Count, 1).End(xlUp).Row


        For i = 2 To RowCount



            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
                Total = Total + ws.Cells(i, 7).Value
                ws.Range("I" & TotalLineNumber).Value = ws.Cells(i, 1).Value
                ws.Range("L" & TotalLineNumber).Value = Total
                TotalLineNumber = TotalLineNumber + 1
                Total = 0
                
                ClosingPrice = ws.Cells(i, 6).Value
                ' Calculate Yearly Change and Percent Change

                    ' ws.Range("M" & TotalLineNumber-1).Value = OpeningPrice
                    ' ws.Range("N" & TotalLineNumber-1).Value = ClosingPrice
                    
                    ws.Range("J" & TotalLineNumber-1).Value = ClosingPrice - OpeningPrice
                    'check for 0 Opening Price
                        If OpeningPrice = 0 then
                        ws.Range("K" & TotalLineNumber-1).Value = "N/A"
                        Else
                        ws.Range("K" & TotalLineNumber-1).Value = FormatPercent((ClosingPrice-OpeningPrice)/OpeningPrice)
                        End if

                    If ws.Range("J" & TotalLineNumber-1).Value >0 Then
                    ws.Range("J" & TotalLineNumber-1).Interior.Colorindex = 4
                    Elseif ws.Range("J" & TotalLineNumber-1).Value <0 Then
                    ws.Range("J" & TotalLineNumber-1).Interior.Colorindex = 3
                    End If 
                ' Reset Opening and closing prices

                OpeningPrice = 0 
                ClosingPrice = 0
            Else

                Total = Total + ws.Cells(i, 7).Value

                If ws.Cells(i, 1).Value <> ws.Cells(i -1, 1).Value Then

                    ' If the Ticker changes, then save the first value
                    OpeningPrice = ws.Cells(i, 3).Value
                    ' Do nothing if the ticker is the same

                End If

            End If
        
        Next i

        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"
        ws.Columns("A:L").AutoFit
        
Next ws
End Sub

