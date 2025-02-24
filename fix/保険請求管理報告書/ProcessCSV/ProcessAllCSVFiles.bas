Sub ProcessAllCSVFiles(fso As Object, newBook As Workbook, csvFolder As String)
    Dim csvFile As Object
    Dim fileType As String
    Dim sheetName As String

    For Each csvFile In fso.GetFolder(csvFolder).Files
        fileType = ""
        sheetName = Left(fso.GetBaseName(csvFile.Name), 30)

        ' CSVの種類を判別して処理
        Select Case True
            Case InStr(csvFile.Name, "fmei") > 0
                fileType = "振込額明細書"
            Case InStr(csvFile.Name, "zogn") > 0
                fileType = "増減点連絡書"
            Case InStr(csvFile.Name, "henr") > 0
                fileType = "返戻内訳書"
            Case Else
                GoTo NextFile
        End Select

        ' CSVデータを転記
        ImportCSVData csvFile.Path, newBook.Sheets.Add, fileType

NextFile:
    Next csvFile
End Sub