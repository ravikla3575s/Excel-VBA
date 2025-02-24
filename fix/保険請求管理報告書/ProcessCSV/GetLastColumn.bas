Function GetLastColumn(ws As Worksheet) As Long
    On Error Resume Next
    GetLastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
    If GetLastColumn = 1 And ws.Cells(1, 1).Value = "" Then GetLastColumn = 0
    On Error GoTo 0
End Function