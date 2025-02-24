Function FindOrCreateReport(savePath As String, csvYYMM As String, templatePath As String) As String
    Dim fso As Object
    Dim filePath As String
    Dim fileName As String
    Dim existingFile As Object
    Dim newWb As Workbook
    
    ' **FileSystemObject の作成**
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' **フォルダ内の報告書 (`RYYMM.xlsx`) を検索**
    For Each existingFile In fso.GetFolder(savePath).Files
        If LCase(fso.GetExtensionName(existingFile.Name)) = "xlsx" Then
            If Right(fso.GetBaseName(existingFile.Name), 4) = csvYYMM Then
                ' **診療年月が一致するファイルが見つかった場合、そのパスを返す**
                FindOrCreateReport = existingFile.Path
                Exit Function
            End If
        End If
    Next existingFile

    ' **該当するファイルがなければ、新規作成**
    fileName = "保険請求管理報告書_R" & csvYYMM & ".xlsx"
    filePath = savePath & "\" & fileName

    ' **テンプレートを元に新規作成**
    Set newWb = Workbooks.Open(templatePath)
    newWb.SaveAs filePath, FileFormat:=xlOpenXMLWorkbook
    newWb.Close
    FindOrCreateReport = filePath
End Function