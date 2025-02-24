Sub ProcessAllCSVFiles(fso As Object, newBook As Workbook, csvFolder As String)
    Dim csvFile As Object
    Dim fileType As String
    Dim wsDetails As Worksheet

    ' シート2（詳細データ用）を取得
    Set wsDetails = newBook.Sheets(2)

    ' CSVファイルをループ処理
    For Each csvFile In fso.GetFolder(csvFolder).Files
        If LCase(fso.GetExtensionName(csvFile.Name)) = "csv" Then
            fileType = ""

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

            ' **請求確定状況の詳細データをシート2に転記**
            TransferBillingDetails newBook, csvFile.Name

NextFile:
        End If
    Next csvFile
End Sub