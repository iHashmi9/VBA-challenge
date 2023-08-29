Attribute VB_Name = "Module5"
Sub GreatestIncDecVolume()
    
    Dim Ticker As String
    Dim TotalStockVolume As Double
    Dim MaxPercent As Double
    Dim MinPercent As Double
    Dim MaxVolume As Double
    MaxPercent = 0
    MinPercent = 0
    MaxVolume = 0
    
    Dim MaxPercentTicker As String
    Dim MinPercentTicker As String
    Dim MaxVolumeTicker As String
    
    Dim SummaryTable As Long
    SummaryTable = 2
    Dim LastRow As Long
    
    For Each ws In Worksheets
        
        LastRow = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
        
        MaxPercent = WorksheetFunction.Max(ws.Range("K:K"))
        ws.Range("Q2").Value = MaxPercent
        
        For i = 2 To LastRow
            If ws.Cells(i, 11).Value = MaxPercent Then
                MaxPercentTicker = ws.Cells(i, 9).Value
                ws.Range("P2").Value = MaxPercentTicker
            End If
        Next i
        
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
        MinPercent = WorksheetFunction.Min(ws.Range("K:K"))
        ws.Range("Q3").Value = MinPercent
        
        For i = 2 To LastRow
            If ws.Cells(i, 11).Value = MinPercent Then
                MinPercentTicker = ws.Cells(i, 9).Value
                ws.Range("P3").Value = MinPercentTicker
            End If
        Next i
        
        ws.Range("Q4").NumberFormat = "0"
        
        MaxVolume = WorksheetFunction.Max(ws.Range("L:L"))
        ws.Range("Q4").Value = MaxVolume
        
        For i = 2 To LastRow
            If ws.Cells(i, 12).Value = MaxVolume Then
                MaxVolumeTicker = ws.Cells(i, 9).Value
                ws.Range("P4").Value = MaxVolumeTicker
            End If
        Next i
        MaxPercent = 0
        MinPercent = 0
        MaxVolume = 0
        ws.Columns("A:R").AutoFit
        
    Next ws
    
End Sub

