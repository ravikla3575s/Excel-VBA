Function GetUniqueSheetName(wb As Workbook, baseName As String) As String
    Dim newName As String
    Dim counter As Integer
    Dim ws As Worksheet
    Dim exists As Boolean

    newName = baseName
    counter = 1

    ' **同じ名前のシートが存在するか確認**
    Do
        exists = False
        For Each ws In wb.Sheets
            If LCase(ws.Name) = LCase(newName) Then
                exists = True
                Exit For
            End If
        Next ws
        If exists Then
            newName = baseName & "_" & counter
            counter = counter + 1
        End If
    Loop While exists

    GetUniqueSheetName = newName
End Function