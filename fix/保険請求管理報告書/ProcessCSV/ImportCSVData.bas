' CSVデータを転記
Sub ImportCSVData(csvFile As String, ws As Worksheet, fileType As String)
    Dim colMap As Object
    Dim fso As Object
    Dim ts As Object
    Dim lineText As String
    Dim dataArray As Variant
    Dim i As Long, j As Long, key
    Dim isHeader As Boolean

    ' エラーハンドリング
    On Error GoTo ErrorHandler

    ' 画面更新・計算を一時停止
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' 項目マッピングを取得
    Set colMap = GetColumnMapping(fileType)

    ' シートをクリア
    ws.Cells.Clear

    ' 1行目に項目名を転記
    j = 1
    For Each key In colMap.Keys
        ws.Cells(1, j).Value = colMap(key)
        j = j + 1
    Next key

    ' CSVデータを読み込んで転記
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(csvFile, 1, False, -2) ' UTF-8対応（-2）

    ' データを転記
    i = 2
    isHeader = True ' 最初の行はヘッダー行
    Do While Not ts.AtEndOfStream
        lineText = ts.ReadLine
        dataArray = Split(lineText, ",")

        ' ヘッダー行をスキップ
        If isHeader Then
            isHeader = False
        Else
            j = 1
            For Each key In colMap.Keys
                If key - 1 <= UBound(dataArray) Then
                    ws.Cells(i, j).Value = Trim(dataArray(key - 1))
                End If
                j = j + 1
            Next key
            i = i + 1
        End If
    Loop
    ts.Close

    ' 列幅を自動調整
    ws.Cells.EntireColumn.AutoFit

    ' 画面更新・計算を再開
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    Exit Sub

' エラーハンドリング
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    If Not ts Is Nothing Then ts.Close
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub