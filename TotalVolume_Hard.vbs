

Sub Totalvolume()



For Each ws In Worksheets

    Dim RowCount As Long
    Dim Total As Double
    Dim TotalLineNumber As Long
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim RowCountGreatest As Long
    Dim GreatestIncrease As Double
    Dim GreatestIncreaseIndex As Integer
    Dim GreastedDescrease As Double
    Dim GreatestDecreaseIndex As Integer
    Dim GreatestTotalVolume As Double
    Dim GreatestTotalVolumeIndex As Integer

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
                    
                    ws.Range("J" & TotalLineNumber - 1).Value = ClosingPrice - OpeningPrice
                    'check for 0 Opening Price
                        If OpeningPrice = 0 Then
                        ws.Range("K" & TotalLineNumber - 1).Value = "N/A"
                        Else
                        ws.Range("K" & TotalLineNumber - 1).Value = FormatPercent((ClosingPrice - OpeningPrice) / OpeningPrice)
                        End If

                    If ws.Range("J" & TotalLineNumber - 1).Value > 0 Then
                    ws.Range("J" & TotalLineNumber - 1).Interior.ColorIndex = 4
                    ElseIf ws.Range("J" & TotalLineNumber - 1).Value < 0 Then
                    ws.Range("J" & TotalLineNumber - 1).Interior.ColorIndex = 3
                    End If
                ' Reset Opening and closing prices

                OpeningPrice = 0
                ClosingPrice = 0
            Else

                Total = Total + ws.Cells(i, 7).Value

                If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then

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

    ' Hard (Greated increase, decrease and total volume)

    ' Greatest Increase
    RowCountGreatest = 0
    RowCountGreatest = ws.Cells(Rows.Count, 11).End(xlUp).Row
    GreatestIncrease = 0
    GreatestIncreaseIndex = 0

    For j = 2 To RowCountGreatest - 1

        If (IsNumeric(ws.Cells(j, 11).Value) = "True") And (ws.Cells(j, 11).Value > GreatestIncrease) Then
        
        GreatestIncrease = ws.Cells(j, 11).Value
        GreatestIncreaseIndex = j
        End If
    
    Next j
    
    ws.Range("Q2").Value = FormatPercent(GreatestIncrease)
    ws.Range("P2").Value = ws.Cells(GreatestIncreaseIndex,9).Value

    ' Greatest Decrease

    RowCountGreatest = 0
    RowCountGreatest = ws.Cells(Rows.Count, 11).End(xlUp).Row
    GreatestDecrease = 0
    GreatestDecreaseIndex = 0

    For j = 2 To RowCountGreatest - 1

        If (IsNumeric(ws.Cells(j, 11).Value) = "True") And (ws.Cells(j, 11).Value < GreatestDecrease) Then
        
        GreatestDecrease = ws.Cells(j, 11).Value
        GreatestDecreaseIndex = j
        End If
    
    Next j
    
    ws.Range("Q3").Value = FormatPercent(GreatestDecrease)
    ws.Range("P3").Value = ws.Cells(GreatestDecreaseIndex,9).Value


    ' Greatest Volume

    RowCountGreatest = 0
    RowCountGreatest = ws.Cells(Rows.Count, 12).End(xlUp).Row
    GreatestTotalVolume = 0
    GreatestTotalVolumeIndex = 0

    For j = 2 To RowCountGreatest - 1

        If (IsNumeric(ws.Cells(j, 12).Value) = "True") And (ws.Cells(j, 12).Value > GreatestTotalVolume) Then
        
        GreatestTotalVolume = ws.Cells(j, 12).Value
        GreatestTotalVolumeIndex = j
        End If
    
    Next j
    
    ws.Range("Q4").Value = GreatestTotalVolume
    ws.Range("P4").Value = ws.Cells(GreatestTotalVolumeIndex,9).Value

' Greatest % Increase
' Greatest % Decrease
' Greatest Total Volume

    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"
    ws.Columns("A:Q").AutoFit

Next ws
End Sub


