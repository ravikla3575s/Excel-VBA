Function FindAllFixfFiles(fso As Object, csvFolder As String) As Object
    Dim fixfFiles As Object
    Dim file As Object
    Set fixfFiles = CreateObject("Scripting.Dictionary")

    For Each file In fso.GetFolder(csvFolder).Files
        If InStr(LCase(file.Name), "fixf") > 0 Then
            fixfFiles.Add file.Path, file
        End If
    Next file
    
    Set FindAllFixfFiles = fixfFiles
End Function