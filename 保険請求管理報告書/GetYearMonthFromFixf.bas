Sub GetYearMonthFromFixf(fixfFile As String, ByRef csvYYMM As String)
    Dim fileName As String
    Dim gyyymm As String
    Dim yearPart As Integer, monthPart As Integer

    ' **ファイル名を取得**
    fileName = Dir(fixfFile) ' フルパスからファイル名のみ取得
    
    ' **GYYMMをファイル名の18～22桁目から取得（不足時は末尾5桁を使用）**
    If Len(fileName) >= 22 Then
        gyyymm = Mid(fileName, 18, 5) ' **18～22桁目を取得**
    Else
        gyyymm = Right(fileName, 5) ' **末尾5桁を取得**
    End If
    
    ' **GYYMM から YYMM へ変換**
    yearPart = CInt(Mid(gyyymm, 2, 2)) ' **2桁の年**
    monthPart = CInt(Right(gyyymm, 2)) ' **2桁の月**

    ' **1月の場合は前年12月に修正**
    If monthPart = 1 Then
        monthPart = 12
        yearPart = yearPart - 1
    Else
        monthPart = monthPart - 1
    End If

    ' **取得した診療年月 (`YYMM`) をセット**
    csvYYMM = Format(yearPart, "00") & Format(monthPart, "00")
End Sub