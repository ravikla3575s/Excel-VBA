Sub TransferBillingDetails(newBook As Workbook, sheetName As String, csvFileName As String)
    Dim wsBilling As Worksheet, wsDetails As Worksheet
    Dim lastRowBilling As Long, lastRowDetails As Long
    Dim i As Long, j As Long
    Dim dispensingMonth As String, convertedMonth As String
    Dim payerCode As String, payerType As String
    Dim receiptNo As String
    Dim startRowDict As Object
    Dim rebillDict As Object, lateDict As Object, unpaidDict As Object, assessmentDict As Object
    Dim category As String
    Dim startRow As Long
    Dim dataDict As Object
    Dim rowData As Variant
    Dim a As Long, b As Long, c As Long

    ' シート設定
    Set wsDetails = newBook.Sheets(2) ' 詳細用シート

    ' CSVデータの請求先分類
    payerCode = Mid(sheetName, 7, 1)
    Select Case payerCode
        Case "1": payerType = "社保"
        Case "2": payerType = "国保"
        Case Else: payerType = "労災"
    End Select

    ' **開始行管理用 Dictionary 作成**
    Set startRowDict = CreateObject("Scripting.Dictionary")
    If payerType = "社保" Then
        startRowDict.Add "返戻再請求", GetStartRow(wsDetails, "社保返戻再請求")
        startRowDict.Add "月遅れ請求", GetStartRow(wsDetails, "社保月遅れ請求")
        startRowDict.Add "返戻・査定", GetStartRow(wsDetails, "社保返戻・査定")
        startRowDict.Add "未請求扱い", GetStartRow(wsDetails, "社保未請求扱い")
    ElseIf payerType = "国保" Then
        startRowDict.Add "返戻再請求", GetStartRow(wsDetails, "国保返戻再請求")
        startRowDict.Add "月遅れ請求", GetStartRow(wsDetails, "国保月遅れ請求")
        startRowDict.Add "返戻・査定", GetStartRow(wsDetails, "国保返戻・査定")
        startRowDict.Add "未請求扱い", GetStartRow(wsDetails, "国保未請求扱い")
    End If

    ' **区分ごとの Dictionary を作成**
    Set rebillDict = CreateObject("Scripting.Dictionary")   ' 返戻再請求
    Set lateDict = CreateObject("Scripting.Dictionary")     ' 月遅れ請求
    Set unpaidDict = CreateObject("Scripting.Dictionary")   ' 未請求扱い
    Set assessmentDict = CreateObject("Scripting.Dictionary") ' 返戻・査定

    lastRowBilling = wsBilling.Cells(Rows.Count, "D").End(xlUp).Row

    ' **請求データを Dictionary に格納**
    For i = 2 To lastRowBilling
        dispensingMonth = wsBilling.Cells(i, 2).Value ' **GYYMM形式の診療月**
        convertedMonth = ConvertToWesternDate(dispensingMonth)
        rowData = Array(wsBilling.Cells(i, 4).Value, convertedMonth, wsBilling.Cells(i, 5).Value, wsBilling.Cells(i, 10).Value)

        ' **現在処理中の診療月（csvYYMM）と異なる場合のみ追加**
        If Right(dispensingMonth, 4) <> csvYYMM Then
            ' **CSVの種類で振り分け**
            If InStr(csvFileName, "fixf") > 0 Then
                ' fixf → ユーザーに選択させる
                If ShowRebillSelectionForm(rowData) Then
                    rebillDict.Add wsBilling.Cells(i, 1).Value, rowData ' **返戻再請求**
                Else
                    lateDict.Add wsBilling.Cells(i, 1).Value, rowData ' **月遅れ請求**
                End If
            ElseIf InStr(csvFileName, "zogn") > 0 Then
                unpaidDict.Add wsBilling.Cells(i, 1).Value, rowData ' **未請求扱い**
            ElseIf InStr(csvFileName, "henr") > 0 Then
                assessmentDict.Add wsBilling.Cells(i, 1).Value, rowData ' **返戻・査定**
            End If
        End If
    Next i

    ' **各カテゴリの追加行数を計算**
    a = 0: b = 0: c = 0
    If rebillDict.Count > 4 Then a = rebillDict.Count - 4
    If lateDict.Count > 4 Then b = lateDict.Count - 4
    If assessmentDict.Count > 4 Then c = assessmentDict.Count - 4

    ' **各転記開始行の調整**
    Dim lateStartRow As Long, assessmentStartRow As Long, unpaidStartRow As Long

    lateStartRow = startRowDict("月遅れ請求") + 1 + a
    assessmentStartRow = startRowDict("返戻・査定") + 1 + a + b
    unpaidStartRow = startRowDict("未請求扱い") + 1 + a + b + c

    ' **行を追加**
    If a + b + c > 0 Then
        wsDetails.Rows(lateStartRow & ":" & lateStartRow + a).Insert Shift:=xlDown
        wsDetails.Rows(assessmentStartRow & ":" & assessmentStartRow + b).Insert Shift:=xlDown
        wsDetails.Rows(unpaidStartRow & ":" & unpaidStartRow + c).Insert Shift:=xlDown
    End If

    ' **各 Dictionary の転記処理（空ならスキップ）**
    If rebillDict.Count > 0 Then
        j = startRowDict("返戻再請求")
        Call TransferData(rebillDict, wsDetails, j, payerType)
    End If

    If lateDict.Count > 0 Then
        j = startRowDict("月遅れ請求")
        Call TransferData(lateDict, wsDetails, j, payerType)
    End If

    If unpaidDict.Count > 0 Then
        j = startRowDict("未請求扱い")
        Call TransferData(unpaidDict, wsDetails, j, payerType)
    End If

    If assessmentDict.Count > 0 Then
        j = startRowDict("返戻・査定")
        Call TransferData(assessmentDict, wsDetails, j, payerType)
    End If

    MsgBox payerType & " のデータ転記が完了しました！", vbInformation, "処理完了"
End Sub