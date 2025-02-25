Sub TransferData(dataDict As Object, ws As Worksheet, startRow As Long, payerType As String)
    Dim key As Variant, rowData As Variant
    Dim j As Long
    Dim payerColumn As Long

    ' **Dictionary が空なら処理しない**
    If dataDict.Count = 0 Then Exit Sub

    ' **payerType に応じた転記列を決定**
    If payerType = "社保" Then
        payerColumn = 8 ' 社保の請求先は H列（8列目）
    ElseIf payerType = "国保" Then
        payerColumn = 9 ' 国保の請求先は I列（9列目）
    Else
        Exit Sub ' 労災の場合は処理しない
    End If

    j = startRow
    For Each key In dataDict.Keys
        rowData = dataDict(key)
        ws.Cells(j, 4).Value = rowData(0) ' 患者氏名
        ws.Cells(j, 5).Value = rowData(1) ' 調剤年月
        ws.Cells(j, 6).Value = rowData(2) ' 医療機関名
        ws.Cells(j, payerColumn).Value = payerType ' 請求先（社保 or 国保）
        ws.Cells(j, 10).Value = rowData(3) ' 請求点数
        j = j + 1
    Next key
End Sub