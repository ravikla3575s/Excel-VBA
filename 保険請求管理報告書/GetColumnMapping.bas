' CSVの種類ごとに項目をマッピング
Function GetColumnMapping(fileType As String) As Object
    Dim colMap As Object
    Set colMap = CreateObject("Scripting.Dictionary")

    Select Case fileType
        Case "振込額明細書"
            colMap.Add 2, "診療（調剤）年月"
            colMap.Add 5, "受付番号"
            colMap.Add 14, "氏名"
            colMap.Add 16, "生年月日"
            colMap.Add 22, "医療保険＿療養の給付＿請求点数"
            colMap.Add 23, "医療保険＿療養の給付＿決定点数"
            colMap.Add 24, "医療保険＿療養の給付＿一部負担金"
            colMap.Add 25, "医療保険＿療養の給付＿金額"
            
            ' 公費データ（第一〜第五）
            Dim i As Integer
            For i = 1 To 5
                colMap.Add 33 + (i - 1) * 10, "第" & i & "公費_請求点数"
                colMap.Add 34 + (i - 1) * 10, "第" & i & "公費_決定点数"
                colMap.Add 35 + (i - 1) * 10, "第" & i & "公費_患者負担金"
                colMap.Add 36 + (i - 1) * 10, "第" & i & "公費_金額"
            Next i

            colMap.Add 82, "算定額合計"

        Case "請求確定状況"
            colMap.Add 4, "診療（調剤）年月"
            colMap.Add 5, "氏名"
            colMap.Add 7, "生年月日"
            colMap.Add 9, "医療機関名称"
            colMap.Add 13, "総合計点数"

            ' 公費データ（第一〜第四）
            For i = 1 To 4
                colMap.Add 16 + (i - 1) * 3, "第" & i & "公費_請求点数"
            Next i

            colMap.Add 30, "請求確定状況"
            colMap.Add 31, "エラー区分"

        Case "増減点連絡書"
            colMap.Add 2, "調剤年月"
            colMap.Add 4, "受付番号"
            colMap.Add 11, "区分"
            colMap.Add 14, "老人減免区分"
            colMap.Add 15, "氏名"
            colMap.Add 21, "増減点数（金額）"
            colMap.Add 22, "事由"

        Case "返戻内訳書"
            colMap.Add 2, "調剤年月(YYMM形式)"
            colMap.Add 3, "受付番号"
            colMap.Add 4, "保険者番号"
            colMap.Add 7, "氏名"
            colMap.Add 9, "請求点数"
            colMap.Add 10, "薬剤一部負担金"
            colMap.Add 12, "一部負担金額"
            colMap.Add 13, "患者負担金額（公費）"
            colMap.Add 14, "事由コード"
    End Select

    Set GetColumnMapping = colMap
End Function