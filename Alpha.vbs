Sub Alphabetical_Testing()


    Dim ws As Worksheet
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim lastRow As Long
    
    
     For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        
        
        ws.Cells(1, 19).Value = "Quarterly Change"
        ws.Cells(1, 20).Value = "Percent Change"
        ws.Cells(1, 21).Value = "Total Volume"
        
        Dim rng As Range
        Dim drng As Range
        Set rng = ws.Range("a:a")
        Set drng = ws.Range("r:r")
        rng.Copy Destination:=drng
        drng.RemoveDuplicates Columns:=Array(1)
        ws.Cells(1, 18).Value = "Ticker"
        
        
        lastRow = ws.Cells(ws.Rows.Count, 18).End(xlUp).Row
        For i = 2 To lastRow
        
        'Total Volume
        ws.Cells(i, 21) = Application.WorksheetFunction.SumIfs(ws.Range("G:G"), ws.Range("A:A"), ws.Cells(i, 18))
    
        'Opening Price
        ws.Cells(i, 22) = Application.WorksheetFunction.VLookup(ws.Cells(i, 18), ws.Range("A:C"), Array(3), False)
        
        'Closing Price
        ws.Cells(i, 23) = Application.WorksheetFunction.VLookup(ws.Cells(i, 18), ws.Range("A:f"), Array(6))
        
        openPrice = Application.WorksheetFunction.VLookup(ws.Cells(i, 18), ws.Range("A:C"), Array(3), False)
        closePrice = Application.WorksheetFunction.VLookup(ws.Cells(i, 18), ws.Range("A:F"), Array(6))
        
        ws.Cells(i, 19) = closePrice - openPrice
        quarterlyChange = closePrice - openPrice
        
        If openPrice <> 0 Then
                    ws.Cells(i, 20) = (quarterlyChange / openPrice)
                    Else
                    ws.Cells(i, 20) = 0
                End If
        ws.Cells(i, 20).NumberFormat = "0.00%"
        
        
        ws.Cells(2, 25).Value = "Greatest % Increase"
        ws.Cells(3, 25).Value = "Greatest % Decrease"
        ws.Cells(4, 25).Value = "Greatest Total Volume"
        ws.Cells(1, 22).Value = "Opening_Price"
        ws.Cells(1, 23).Value = "Closing_Price"
        
        
        ws.Cells(2, 27).Value = Application.WorksheetFunction.Max(ws.Range("T:T"))
        ws.Cells(3, 27).Value = Application.WorksheetFunction.Min(ws.Range("T:T"))
        ws.Cells(4, 27).Value = Application.WorksheetFunction.Max(ws.Range("U:U"))
        ws.Cells(2, 26).Value = Application.WorksheetFunction.Index(ws.Range("R:R"), WorksheetFunction.Match(ws.Cells(2, 27).Value, ws.Range("T:T"), 0))
        ws.Cells(3, 26).Value = Application.WorksheetFunction.Index(ws.Range("R:R"), WorksheetFunction.Match(ws.Cells(3, 27).Value, ws.Range("T:T"), 0))
        ws.Cells(4, 26).Value = Application.WorksheetFunction.Index(ws.Range("R:R"), WorksheetFunction.Match(ws.Cells(4, 27).Value, ws.Range("U:U"), 0))


        
        ws.Cells(i, 27).NumberFormat = "0.00%"
        ws.Cells(4, 27).NumberFormat = "#,##0.00"
        ws.Cells(i, 21).NumberFormat = "#,##0.00"
        
        If quarterlyChange > 0 Then
         ws.Cells(i, 19).Interior.ColorIndex = 4
         ElseIf quarterlyChange < 0 Then
         ws.Cells(i, 19).Interior.ColorIndex = 3
         End If
              
        
        Next i
        
    Next ws
        
    End Sub
