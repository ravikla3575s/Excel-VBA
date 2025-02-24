Function GetLastRow(ws As Worksheet) As Long
    On Error Resume Next
    GetLastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    If GetLastRow = 1 And ws.Cells(1, 1).Value = "" Then GetLastRow = 0
    On Error GoTo 0
End Function