Sub TransferData(dataDict As Object, ws As Worksheet, startRow As Long, payerType As String)
    Dim key As Variant, rowData As Variant
    Dim j As Long

    ' **Dictionary が空なら処理しない**
    If dataDict.Count = 0 Then Exit Sub

    j = startRow
    For Each key In dataDict.Keys
        rowData = dataDict(key)
        ws.Cells(j, 4).Value = rowData(0) ' 患者氏名
        ws.Cells(j, 5).Value = rowData(1) ' 調剤年月
        ws.Cells(j, 6).Value = rowData(2) ' 医療機関名
        ws.Cells(j, 8).Value = payerType ' 請求先
        ws.Cells(j, 10).Value = rowData(3) ' 請求点数
        j = j + 1
    Next key
End Sub