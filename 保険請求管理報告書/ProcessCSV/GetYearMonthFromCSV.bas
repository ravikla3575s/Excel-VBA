Sub GetYearMonthFromCSV(fso As Object, csvFolder As String, ByRef targetYear As String, ByRef targetMonth As String)
    Dim csvFile As Object
    Dim ts As Object
    Dim lineText As String
    Dim yearPart As String, monthPart As String
    Dim era As String, westernYear As Integer
    
    ' フォルダ内の最初の CSV を探す
    For Each csvFile In fso.GetFolder(csvFolder).Files
        If LCase(Right(csvFile.Name, 4)) = ".csv" Then
            ' CSV を開く
            Set ts = fso.OpenTextFile(csvFile.Path, 1, False, -2) ' UTF-8対応

            ' 1行目または2行目から GYYMM を取得（例: "50406" = 令和4年6月）
            Do While Not ts.AtEndOfStream
                lineText = ts.ReadLine
                If Len(lineText) >= 5 Then
                    era = Left(lineText, 1) ' 和暦の元号 (1: 明治, 2: 大正, 3: 昭和, 4: 平成, 5: 令和)
                    yearPart = Mid(lineText, 2, 2) ' 2桁の年
                    monthPart = Right(lineText, 2) ' 2桁の月

                    ' 和暦を西暦に変換
                    Select Case era
                        Case "1": westernYear = 1867 + CInt(yearPart) ' 明治
                        Case "2": westernYear = 1911 + CInt(yearPart) ' 大正
                        Case "3": westernYear = 1925 + CInt(yearPart) ' 昭和
                        Case "4": westernYear = 1988 + CInt(yearPart) ' 平成
                        Case "5": westernYear = 2018 + CInt(yearPart) ' 令和
                        Case Else: westernYear = 2000 ' 不明な場合は適当なデフォルト値
                    End Select

                    ' 取得した年月をセット
                    targetYear = CStr(westernYear)
                    targetMonth = monthPart
                    Exit Sub
                End If
            Loop

            ' ファイルを閉じる
            ts.Close
        End If
    Next csvFile
End Sub