Sub GetYearMonthFromFixf(fixfFile As String, ByRef targetYear As String, ByRef targetMonth As String)
    Dim fileName As String
    Dim gyymm As String
    Dim yearPart As Integer, monthPart As Integer
    Dim era As String
    Dim westernYear As Integer

    ' **ファイル名を取得**
    fileName = Dir(fixfFile) ' フルパスからファイル名のみ取得
    
    ' **GYYMMをファイル名の18～22桁目から取得（不足時は末尾5桁を使用）**
    If Len(fileName) >= 22 Then
        gyymm = Mid(fileName, 18, 5) ' **18～22桁目を取得**
    Else
        gyymm = Right(fileName, 5) ' **末尾5桁を取得**
    End If

    ' **和暦の元号を取得（1桁目）**
    era = Left(gyymm, 1) ' **例: 5（令和）**
    
    ' **GYYMM から YYMM へ変換**
    yearPart = CInt(Mid(gyymm, 2, 2)) ' **2桁の年**
    monthPart = CInt(Right(gyymm, 2)) ' **2桁の月**

    ' **1月の場合は前年12月に修正**
    If monthPart = 1 Then
        monthPart = 12
        yearPart = yearPart - 1
    Else
        monthPart = monthPart - 1
    End If

    ' **和暦を西暦に変換**
    Select Case era
        Case "1": westernYear = 1867 + yearPart ' 明治
        Case "2": westernYear = 1911 + yearPart ' 大正
        Case "3": westernYear = 1925 + yearPart ' 昭和
        Case "4": westernYear = 1988 + yearPart ' 平成
        Case "5": westernYear = 2018 + yearPart ' 令和
        Case Else: westernYear = 2000 ' 不明な場合はデフォルト値
    End Select

    ' **取得した診療年月をセット**
    targetYear = CStr(westernYear)
    targetMonth = Format(monthPart, "00")
End Sub