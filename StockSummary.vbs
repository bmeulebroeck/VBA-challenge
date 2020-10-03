Attribute VB_Name = "StockSummary"
Sub StockSummary()
    
    
    For Each ws In Worksheets
    
        Dim sheetname As String
        sheetname = ws.Name
        
        Dim row As Long
        Dim tickerlist As String
        Dim yearopen As Double
        Dim yearclose As Double
        Dim totalvol As Variant
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"
        
        Dim summarytablerow As Integer
        summarytablerow = 2
        yearopen = 0
        
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row
        
        For row = 2 To lastrow
            Do Until yearopen <> 0
                yearopen = ws.Cells(row, 3).Value
                If yearopen = 0 Then Exit Do
            Loop
            
            If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
                tickerlist = ws.Cells(row, 1).Value
                yearclose = ws.Cells(row, 6).Value
                totalvol = totalvol + ws.Cells(row, 7).Value
                ws.Range("I" & summarytablerow).Value = tickerlist
                ws.Range("J" & summarytablerow).Value = (yearclose - yearopen)
                    If ws.Range("J" & summarytablerow).Value >= 0 Then
                        ws.Range("J" & summarytablerow).Interior.ColorIndex = 4
                    Else
                        ws.Range("J" & summarytablerow).Interior.ColorIndex = 3
                    End If
                ws.Range("K" & summarytablerow).Value = (yearclose / yearopen) - 1
                On Error Resume Next
                ws.Range("K" & summarytablerow).NumberFormat = "0.00%"
                ws.Range("L" & summarytablerow).Value = totalvol
                summarytablerow = summarytablerow + 1
                yearopen = 0
                yearclose = 0
                totalvol = 0
            Else
                totalvol = totalvol + ws.Cells(row, 7).Value
                yearclose = 0
            End If
        Next row
        
        Dim pctinc As Double
        Dim pctdec As Double
        Dim highestvol As Variant
        
        pctinc = 0
        pctdec = 0
        highestvol = 0
        
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        
        For row = 2 To lastrow
            If ws.Cells(row, 11).Value > pctinc Then
                pctinc = ws.Cells(row, 11).Value
                ws.Cells(2, 15).Value = ws.Cells(row, 9).Value
                ws.Cells(2, 16).Value = pctinc
                ws.Cells(2, 16).NumberFormat = "0.00%"
            End If
            If ws.Cells(row, 11).Value < pctdec Then
                pctdec = ws.Cells(row, 11).Value
                ws.Cells(3, 15).Value = ws.Cells(row, 9).Value
                ws.Cells(3, 16).Value = pctdec
                ws.Cells(3, 16).NumberFormat = "0.00%"
            End If
            If ws.Cells(row, 12).Value > highestvol Then
                highestvol = ws.Cells(row, 12).Value
                ws.Cells(4, 15).Value = ws.Cells(row, 9).Value
                ws.Cells(4, 16).Value = highestvol
            End If
        Next row
        
        MsgBox ("Sheet complete")
        
    Next ws
    
    MsgBox ("Summary by sheet complete")
    
End Sub
