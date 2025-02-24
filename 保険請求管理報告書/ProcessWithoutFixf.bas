Sub ProcessWithoutFixf(fso As Object, csvFolder As String, savePath As String, templatePath As String)
    Dim targetYear As String
    Dim targetMonth As String
    Dim targetFile As String
    Dim newBook As Workbook

    ' 【1】対象年月を取得
    Call GetYearMonthFromCSV(fso, csvFolder, targetYear, targetMonth)

    ' 【2】対象Excelファイルを取得 or 作成
    targetFile = FindOrCreateReport(savePath, targetYear, targetMonth, templatePath)

    ' 【3】Excelを開き、テンプレート情報を設定
    Set newBook = Workbooks.Open(targetFile)
    SetTemplateInfo(newBook, targetYear, targetMonth)

    ' 【4】フォルダ内のすべてのCSVを処理
    ProcessAllCSVFiles(fso, newBook, csvFolder)

    ' 【5】保存して閉じる
    newBook.Save
    newBook.Close
End Sub