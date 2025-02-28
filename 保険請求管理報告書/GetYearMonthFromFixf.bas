Sub GetYearMonthFromFixf(fixfFile As String, ByRef targetYear As String, ByRef targetMonth As String)
    Dim fileName As String
    Dim datePart As String
    Dim yearPart As String, monthPart As String, dayPart As String
    Dim hourPart As String, minPart As String, secPart As String
    
    ' 【1】ファイル名を取得（フォルダパスを除く）
    fileName = Mid(fixfFile, InStrRev(fixfFile, "\") + 1)

    ' 【2】fixfファイルの日付部分を取得（後半部分）
    ' 例: RTfixf1014123456720250228150730.csv → "20250228150730"
    datePart = Mid(fileName, 18, 14)

    ' 【3】年月日を分解
    yearPart = Left(datePart, 4)    ' "2025"
    monthPart = Mid(datePart, 5, 2) ' "02"
    dayPart = Mid(datePart, 7, 2)   ' "28"
    hourPart = Mid(datePart, 9, 2)  ' "15"
    minPart = Mid(datePart, 11, 2)  ' "07"
    secPart = Mid(datePart, 13, 2)  ' "30"

    ' 【4】取得した年と月を戻り値に設定
    targetYear = yearPart  ' 西暦のまま
    targetMonth = monthPart ' 2桁の月

    ' **確認用ログ（必要なら表示）**
    ' MsgBox "診療年月取得: " & targetYear & "年 " & targetMonth & "月 " & dayPart & "日 " & hourPart & ":" & minPart & ":" & secPart, vbInformation, "確認"
End Sub