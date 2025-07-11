

Sub FormatSheet()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    i = 1
    
    Do While i <= lastRow
        If Trim(ws.Cells(i, 1).Value) Like "GM*" Then
            ' Start of a new table
            Call ColorHeader(ws, i)         ' Row 1 - docId
            Call AddBorders(ws, i)
            
            If i + 1 <= lastRow Then
                Call ColorHeader(ws, i + 1) ' Row 2
                Call AddBorders(ws, i + 1)
            End If
            
            If i + 2 <= lastRow Then
                Call ColorHeader(ws, i + 2) ' Row 3
                Call AddBorders(ws, i + 2)
            End If
            
            If i + 3 <= lastRow Then
                Call ColorDataRow(ws, i + 3) ' Row 4
                Call AddBorders(ws, i + 3)
            End If
            
            ' Remaining rows → default fill, just add borders
            Dim j As Long
            j = i + 4
            Do While j <= lastRow And Trim(ws.Cells(j, 1).Value) <> ""
                Call AddBorders(ws, j)
                j = j + 1
            Loop
            
            i = j ' Skip to next block
        Else
            i = i + 1
        End If
    Loop
End Sub

Sub ColorHeader(ws As Worksheet, rowNum As Long)
    ws.Range("A" & rowNum & ":J" & rowNum).Interior.Color = RGB(255, 235, 156) ' Orange Accent 2, Lighter 80%
End Sub

Sub ColorDataRow(ws As Worksheet, rowNum As Long)
    ws.Range("A" & rowNum & ":J" & rowNum).Interior.Color = RGB(221, 235, 247) ' Blue Accent 1, Lighter 60%
End Sub

Sub AddBorders(ws As Worksheet, rowNum As Long)
    Dim rng As Range
    Set rng = ws.Range("A" & rowNum & ":J" & rowNum)
    
    With rng.Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlThin
    End With
End Sub
