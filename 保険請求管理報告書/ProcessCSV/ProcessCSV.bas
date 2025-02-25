Sub ProcessCSV()
    Dim csvFolder As String
    Dim fso As Object
    Dim targetYear As String
    Dim targetMonth As String
    Dim savePath As String
    Dim templatePath As String
    Dim newBook As Workbook
    Dim targetFile As String
    Dim fixfFile As String
    Dim fixfFiles As Object
    Dim file As Object
    
    ' 【1】CSVフォルダをユーザーに選択させる
    csvFolder = SelectCSVFolder()
    If csvFolder = "" Then Exit Sub

    ' 【2】テンプレートパス・保存フォルダ取得
    templatePath = GetTemplatePath()
    savePath = GetSavePath()
    If templatePath = "" Or savePath = "" Then Exit Sub

    ' 【3】ファイルシステムオブジェクトの作成
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 【4】フォルダ内のすべての `fixf` ファイルを取得
    Set fixfFiles = FindAllFixfFiles(fso, csvFolder)
    
    ' 【5】`fixf` ファイルが存在しない場合、通常のCSV処理に切り替え
    If fixfFiles.Count = 0 Then
        Call ProcessWithoutFixf(fso, csvFolder, savePath, templatePath)
        Exit Sub
    End If

    ' 【6】複数の `fixf` を順番に処理
    For Each file In fixfFiles
        fixfFile = file.Path
        
        ' 【7】対象年月を取得
        Call GetYearMonthFromFixf(fixfFile, targetYear, targetMonth)
        
        ' 【8】対象Excelファイルを取得 or 作成
        targetFile = FindOrCreateReport(savePath, targetYear, targetMonth, templatePath)
        
        ' 【9】Excelを開き、テンプレート情報を設定
        Set newBook = Workbooks.Open(targetFile)
        SetTemplateInfo(newBook, targetYear, targetMonth)
        
        ' 【10】フォルダ内のすべてのCSVを処理
        ProcessAllCSVFiles(fso, newBook, csvFolder)
        
        ' 【11】保存して閉じる
        newBook.Save
        newBook.Close
    Next file

    ' 【12】処理完了メッセージ
    MsgBox "すべての `fixf` ファイルを処理しました！", vbInformation
End Sub