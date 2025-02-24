Sub ProcessCSV()
    Dim csvFolder As String
    Dim fso As Object
    Dim targetYear As String, targetMonth As String
    Dim savePath As String, templatePath As String
    Dim newBook As Workbook
    Dim targetFile As String
    
    ' 【1】CSVフォルダをユーザーに選択させる
    csvFolder = SelectCSVFolder()
    If csvFolder = "" Then Exit Sub

    ' 【2】テンプレートパス・保存フォルダ取得
    templatePath = GetTemplatePath()
    savePath = GetSavePath()
    If templatePath = "" Or savePath = "" Then Exit Sub

    ' 【3】ファイルシステムオブジェクトの作成
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 【4】fixfファイルの有無をチェック（なくても処理可能）
    Dim fixfFile As String
    fixfFile = FindFixfFile(fso, csvFolder)

    ' 【5】対象年月の取得（fixfがなくても処理できるように）
    If fixfFile <> "" Then
        GetYearMonthFromFixf fixfFile, targetYear, targetMonth
    Else
        GetYearMonthFromCSV fso, csvFolder, targetYear, targetMonth
    End If

    ' 【6】対象Excelファイルを取得 or 作成
    targetFile = FindOrCreateReport(savePath, targetYear, targetMonth, templatePath)

    ' 【7】Excelを開き、テンプレート情報を設定
    Set newBook = Workbooks.Open(targetFile)
    SetTemplateInfo newBook, targetYear, targetMonth

    ' 【8】フォルダ内のすべてのCSVを処理
    ProcessAllCSVFiles fso, newBook, csvFolder)

    ' 【9】保存して閉じる
    newBook.Save
    newBook.Close
    MsgBox "すべてのCSVを処理しました！", vbInformation
End Sub