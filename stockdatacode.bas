Attribute VB_Name = "Module1"
Sub test():
    For Each ws In Worksheets
        Dim WorksheetName As String
        WorksheetName = ws.Name
        Dim i As Long
        Dim j As Long
        j = 2
        Dim ticker As String
        Dim year_change As Double
        Dim per_change As Double
        Dim volume As Long
        volume = 0
        Dim row As Long
        row = 2
        Dim greatvol As Double
        Dim greatincr As Double
        Dim greatdecr As Double
        Dim lastrowA As Long
        Dim lastrowI As Long
        lastrowA = ws.Cells(Rows.Count, 1).End(xlUp).row
            For i = 2 To lastrowA
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ws.Cells(row, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(row, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                    If ws.Cells(row, 10).Value < 0 Then
                    ws.Cells(row, 10).Interior.ColorIndex = 3
                    Else
                    ws.Cells(row, 10).Interior.ColorIndex = 4
                    End If
                    If ws.Cells(j, 3).Value <> 0 Then
                    per_change = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    ws.Cells(row, 11).Value = Format(per_change, "Percent")
                    Else
                    ws.Cells(row, 11).Value = Format(0, "Percent")
                    End If
                ws.Cells(row, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                row = row + 1
                j = i + 1
                End If
            Next i
        lastrowI = ws.Cells(Rows.Count, 9).End(xlUp).row
        greatvol = ws.Cells(2, 12).Value
        greatincr = ws.Cells(2, 11).Value
        greatdecr = ws.Cells(2, 11).Value
            For i = 2 To lastrowI
                If ws.Cells(i, 12).Value > greatvol Then
                greatvol = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                Else
                greatvol = greatvol
                End If
                If ws.Cells(i, 11).Value > greatincr Then
                greatincr = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                Else
                greatincr = greatincr
                End If
                If ws.Cells(i, 11).Value < greatdecr Then
                greatdecr = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                Else
                greatdecr = greatdecr
                End If
            ws.Cells(2, 17).Value = Format(greatincr, "Percent")
            ws.Cells(3, 17).Value = Format(greatdecr, "Percent")
            ws.Cells(4, 17).Value = Format(greatvol, "Scientific")
            Next i
            
    Next ws
End Sub
