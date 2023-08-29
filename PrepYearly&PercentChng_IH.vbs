Attribute VB_Name = "Module3"
Sub YearlyChange_PercentChange_MaxMin()

For Each ws In Worksheets

    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    OpenPrice = 0
    ClosePrice = 0
    YearlyChange = 0
    PercentChange = 0
    
    Dim SummaryTable As Long
    SummaryTable = 2
    Dim LastRow As Long
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        OpenPrice = ws.Cells(2, 3).Value
            For i = 2 To LastRow
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    Ticker = ws.Cells(i, 1).Value
                    
                    ClosePrice = ws.Cells(i, 6).Value
                    YearlyChange = ClosePrice - OpenPrice
                    
                If OpenPrice <> 0 Then
                    PercentChange = (YearlyChange / OpenPrice) * 100
                End If
                
            ws.Range("J" & SummaryTable).Value = YearlyChange
                If (YearlyChange > 0) Then
                    ws.Range("J" & SummaryTable).Interior.ColorIndex = 4
                        ElseIf (YearlyChange <= 0) Then
                    ws.Range("J" & SummaryTable).Interior.ColorIndex = 3
                End If
        
                ws.Range("K" & SummaryTable).Value = (CStr(PercentChange) & "%")
            
                SummaryTable = SummaryTable + 1
        
                YearlyChange = 0
                ClosePrice = 0
        
                OpenPrice = ws.Cells(i + 1, 3).Value
                
                End If

            Next i
        
        Next ws

End Sub
