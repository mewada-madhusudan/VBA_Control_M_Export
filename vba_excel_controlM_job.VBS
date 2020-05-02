Function ActivateWB(wbname As String)
  'This function takes a workbook name as string as input and opens
  'Open wbname.

  Workbooks(wbname).Activate

End Function
Function LastRow(Sh As Worksheet)
'This Fucntion takes a worksheet as an input and returns the last used row in the sheet

    On Error Resume Next
    LastRow = Sh.Cells.Find(What:="*", _
                            After:=Sh.Range("A1"), _
                            Lookat:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Row
    On Error GoTo 0
End Function

Function LastCol(Sh As Worksheet)
'This Fucntion takes a worksheet as an input and returns the last used row in the sheet
'SearchOrder:=xlByRows, _


    On Error Resume Next
    LastCol = Sh.Cells.Find(What:="*", _
                            After:=Sh.Range("A1"), _
                            Lookat:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Column
    On Error GoTo 0
End Function

Function findVal(Val As String)
'This Fucntion takes a worksheet as an input and returns the last used row in the sheet

Dim CheckCol As Range
Set CheckCol = ActiveSheet.Range("A:A")
MsgBox (CheckCol)

    On Error Resume Next
    findVal = CheckCol.Find(What:=Val, _
                            After:=Sh.Range("A1"), _
                            Lookat:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Row
    On Error GoTo 0
End Function


Sub GetData()
With Application
    .ScreenUpdating = False
End With


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Declare Variables                                                  '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Val As String
Dim Srcbk As Workbook
Dim SrcSh As Worksheet
Dim DestSh As Worksheet
Dim rngX As Range
Dim name As Integer
Dim week As Integer
Dim colnum As Integer 'this is a counter used to get the row to read from the src sheet
Dim SrcWkb As Workbook

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Set Variables                                                      '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Set Destination Sheet
Set DestSh = ActiveWorkbook.Worksheets("Sheet1")

'Source File name
StrSheet = "state.xlsx"
'Source File Path
strPath = "C:\Users\madhu\Desktop\" & StrSheet
'MsgBox (strPath)
'set Source Workbook
Set xSourceWb = Workbooks.Open(strPath)
'Set SrcWkb = appExcel.ActiveWorkbook
'Set Source Sheet
Set SrcSh = xSourceWb.Worksheets.Item(1)

'Set SrcSh = lookUpBook.Worksheets("Resource Detailed Report")
'Search Range
'Set rngX = SrcSh.Range("B:B").Find(Val, lookat:=xlPart)


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Code Execution                                                     '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'find last cell in destination sheet
    last = LastRow(DestSh)
'find last cell in destination sheet - this should return the total number of weeks data is avaiable for
    endofpage = LastCol(SrcSh) - 4
    
    
    

'Loop Through Each name in the lookup sheet starts at row 3 for now
    For name = 2 To last
    
    'Define Value you want to look up
    Val = DestSh.Cells(name, 1).Value
   ' MsgBox (Val)
    If Not Val = "" Then
    MsgBox (Val)
            'Search Range
            Set rngX = SrcSh.Range("A:B").Find(Val, Lookat:=xlPart)
            MsgBox (rngX.Address)
           ' MsgBox (rngX.Value)
            For destCol = 2 To 2
                MsgBox (destCol)
                If Not rngX Is Nothing Then
                    'Compliance
                    'DestSh.Cells(name, destCol).Value = rngX.Value
                   ' MsgBox (SrcSh.Cells(name, destCol).Value)
                    'Sheet1.Cells(name, 2) = Cells(rngX.Row, 2).Value
                    MsgBox (Sheet1.Cells(name, 2))
                    MsgBox (Cells(rngX.Row, 2).Value)
                    If Cells(rngX.Row, 2).Value = "Ended OK" Then
                    DestSh.Cells(name, 1).Interior.ColorIndex = 4
                    End If
                    If Cells(rngX.Row, 2).Value = "Ended Not OK" Then
                    DestSh.Cells(name, 1).Interior.ColorIndex = 3
                    End If
                    If Cells(rngX.Row, 2).Value = "Executing" Then
                    DestSh.Cells(name, 1).Interior.ColorIndex = 6
                    End If
                    If Cells(rngX.Row, 2).Value = "Waiting Condition" Then
                    DestSh.Cells(name, 1).Interior.ColorIndex = 17
                    End If
                Else
                    MsgBox ("Cannot retrieve data for " & Val & ". Name was not found in " & SrcSh.name)
                    GoTo NextName
                End If
            Next
    End If
    
NextName:
    Next
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Cleanup                                                            '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Cleanup:

'Workbooks("C:\Users\ctwellma\SkyDrive\datepicker.xls").Worksheets("Sheet1").Activate

With Application
    .ScreenUpdating = True
End With
     
End Sub

