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

Sub mycro()

Dim Val As String
Dim Srcbk As Workbook
Dim SrcSh As Worksheet
Dim DestSh As Worksheet
Dim rngX As Range
Dim name As Integer
Dim week As Integer
Dim colnum As Integer 'this is a counter used to get the row to read from the src sheet
Dim SrcWkb As Workbook
Dim Tdate As Date
Dim Val1, Val2 As String


Tdate = Date
'MsgBox (Tdate - 1)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Set Variables                                                      '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Set Destination Sheet
Set DestSh = ActiveWorkbook.Worksheets("Sheet1")
'Source File name
StrSheet = "states.csv"
'Source File Path
strPath = "C:\Users\madhu\Desktop\" & StrSheet
Set xSourceWb = Workbooks.Open(strPath)
Set SrcSh = xSourceWb.Worksheets.Item(1)
'find last cell in destination sheet
    last = LastRow(DestSh)
'find last cell in destination sheet - this should return the total number of weeks data is avaiable for
    endofpage = LastRow(SrcSh)
    
    
'Loop Through Each name in the lookup sheet starts at row 3 for now
    'For name = 2 To last
    
    'Define Value you want to look up
    'Val = DestSh.Cells(name, 1).Value
    'MsgBox (Val)
    'If Val = Tdate - 1 Then
     
        
    'End If
    
    
    
    For name = 2 To last + 1
        MsgBox (name)
             'Set rngX = SrcSh.Range("A:B").Find(DestSh.Cells(name, 2).Value, Lookat:=xlPart)
             Set rngX = SrcSh.Range("A:B").Find(Tdate, Lookat:=xlPart)
            MsgBox (rngX.Address)
            Set rngX = SrcSh.Range("A:B").Find(Tdate - 1, Lookat:=xlPart)
            MsgBox (rngX.Address)
                If DestSh.Cells(name, 1).Value = Tdate - 1 And SrcSh.Cells(rngX.Row, 2).Value = DestSh.Cells(name, 2).Value Then
                    'MsgBox (rngX.Row)
                    If SrcSh.Cells(rngX.Row, 2).Value = DestSh.Cells(name, 2).Value Then
                        DestSh.Cells(name, 3).Value = SrcSh.Cells(rngX.Row, 3).Value
                    End If
                    
                End If
            
    Next
    
    
    

    

With Application
    .ScreenUpdating = True
End With




End Sub


