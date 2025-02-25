Function FindAllFixfFiles(fso As Object, csvFolder As String) As Collection
    Dim csvFile As Object
    Dim fixfFiles As New Collection

    ' **フォルダ内のすべてのファイルをループ**
    For Each csvFile In fso.GetFolder(csvFolder).Files
        ' **拡張子が "csv" であり、名前に "fixf" を含む場合**
        If LCase(fso.GetExtensionName(csvFile.Name)) = "csv" And InStr(LCase(csvFile.Name), "fixf") > 0 Then
            fixfFiles.Add csvFile ' **fixf ファイルをコレクションに追加**
        End If
    Next csvFile

    ' **コレクションを返す**
    Set FindAllFixfFiles = fixfFiles
End Function