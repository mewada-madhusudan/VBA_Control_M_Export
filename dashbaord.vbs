Private Sub Worksheet_Change(ByVal Target As Range)
WorkbookChange Target
End Sub



Private Sub Workbook_Open()

MsgBox "Welcome back, I missed you!"


End Sub


Sub WorkbookChange(ByVal Target As Range)
Dim cell As Range
Set rng = Range("Sheet2!A2:A6")
For Each cell In rng
    If cell.Interior.ColorIndex = 3 Then
        'MsgBox (cell.Interior.ColorIndex)
        Range("Sheet1!A1").Interior.ColorIndex = 17
    ElseIf cell.Interior.ColorIndex = 6 Then
        temp = 6
    ElseIf cell.Interior.ColorIndex = 17 Then
        temp = 4
    ElseIf cell.Interior.ColorIndex = 4 Then
        temp = 17
    End If    
Next cell
Goto finish

'Set rng = Range("Sheet2!B2:B6")
'For Each cell In rng
    
 '   Range("Sheet1!B1").Interior.ColorIndex = 4
    
'Next cell

End Sub

