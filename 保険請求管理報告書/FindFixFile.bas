Function FindFixfFile(fso As Object, csvFolder As String) As String
    Dim csvFile As Object
    For Each csvFile In fso.GetFolder(csvFolder).Files
        If InStr(LCase(csvFile.Name), "fixf") > 0 Then
            FindFixfFile = csvFile.Path
            Exit Function
        End If
    Next csvFile
    FindFixfFile = "" ' fixfファイルがない場合
End Function