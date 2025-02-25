Sub ProcessAllCSVFiles(fso As Object, newBook As Workbook, csvFolder As String)
    Dim csvFile As Object
    Dim fileType As String
    Dim wsDetails As Worksheet
    Dim wsCSV As Worksheet
    Dim sheetName As String
    Dim sheetIndex As Integer

    ' シート2（詳細データ用）を取得
    Set wsDetails = newBook.Sheets(2)

    ' CSVファイルをループ処理
    For Each csvFile In fso.GetFolder(csvFolder).Files
        If LCase(fso.GetExtensionName(csvFile.Name)) = "csv" Then
            fileType = ""

            ' **CSVの種類を判別**
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

            ' **シート名にファイル名（拡張子なし）を設定**
            sheetName = fso.GetBaseName(csvFile.Name)

            ' **シートがすでに存在する場合、"_1", "_2" を付けて回避**
            sheetName = GetUniqueSheetName(newBook, sheetName)

            ' **シートを3番目（Sheets(3)の位置）に追加**
            sheetIndex = Application.WorksheetFunction.Min(3, newBook.Sheets.Count + 1)
            Set wsCSV = newBook.Sheets.Add(After:=newBook.Sheets(sheetIndex - 1))
            wsCSV.Name = sheetName

            ' **CSVデータを転記**
            ImportCSVData(csvFile.Path, wsCSV, fileType)

            ' **請求確定状況の詳細データをシート2に転記**
            TransferBillingDetails(newBook, csvFile.Name)

NextFile:
        End If
    Next csvFile
End Sub