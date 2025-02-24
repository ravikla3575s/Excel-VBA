Function FindOrCreateReport(savePath As String, targetYear As String, targetMonth As String, templatePath As String) As String
    Dim fso As Object
    Dim filePath As String
    Dim fileName As String
    Dim newWb As Workbook

    ' **FileSystemObject を作成**
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' **保存先のファイル名を決定**
    fileName = "保険請求管理報告書_" & targetYear & targetMonth & ".xlsx"
    filePath = savePath & "\" & fileName

    ' **既存ファイルがあるかチェック**
    If fso.FileExists(filePath) Then
        FindOrCreateReport = filePath
    Else
        ' **テンプレートを元に新規作成**
        Set newWb = Workbooks.Open(templatePath)
        newWb.SaveAs filePath, FileFormat:=xlOpenXMLWorkbook
        newWb.Close
        FindOrCreateReport = filePath
    End If
End Function