Attribute VB_Name = "Module2"
Sub Preparing_Total_Stock()

    Dim Ticker As String
    Dim TotalStockVolume As Double
    Dim LastRow As Double
    Dim SummaryTable As Long
    Dim CurrentRow As Long 
    
    SummaryTable = 2
        
    For Each ws In Worksheets
    
        LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        LastRowB = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
                        
        For i = 2 To LastRowB
            ws.Cells(i, 2).Value = CDbl(ws.Cells(i, 2).Value)
        Next i
                            
        For j = 2 To LastRow
            If ws.Cells(j, 1).Value <> ws.Cells(j + 1, 1).Value Then
                
                Ticker = ws.Cells(j, 1).Value
                TotalStockVolume = TotalStockVolume + ws.Cells(j, 7).Value
                             
                ws.Range("I" & SummaryTable).Value = Ticker 
                ws.Range("L" & SummaryTable).Value = TotalStockVolume 
            
                TotalStockVolume = 0
                            
                SummaryTable = SummaryTable + 1 
                
            Else
            
                TotalStockVolume = TotalStockVolume + ws.Cells(j, 7).Value
                
            End If
       
        Next j
    
        SummaryTable = 2
        
    Next ws

End Sub



