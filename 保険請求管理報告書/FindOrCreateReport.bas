Function FindOrCreateReport(savePath As String, targetYear As String, targetMonth As String, templatePath As String) As String
    Dim fso As Object
    Dim filePath As String
    Dim fileName As String
    Dim existingFile As Object
    Dim newWb As Workbook
    Dim csvYYMM As String
    Dim sheet1Name As String, sheet2Name As String
    
    ' **診療年月を "YYMM" 形式に変換**
    csvYYMM = Right(targetYear, 2) & targetMonth

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
    
    ' **シート名を診療年月に応じて変更**
    sheet1Name = "R" & CInt(targetYear) & "." & CInt(targetMonth) ' R y.m 形式
    sheet2Name = ConvertToCircledNumber(CInt(targetMonth)) ' ①～⑫ に変換

    On Error Resume Next
    newWb.Sheets(1).Name = sheet1Name
    newWb.Sheets(2).Name = sheet2Name
    On Error GoTo 0
    
    ' **ブックを閉じてパスを返す**
    newWb.Close SaveChanges:=True
    FindOrCreateReport = filePath
End Function