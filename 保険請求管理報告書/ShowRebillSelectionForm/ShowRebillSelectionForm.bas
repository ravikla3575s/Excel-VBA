Sub ShowRebillSelectionForm(newBook As Workbook)
    Dim wsBilling As Worksheet
    Dim lastRow As Long, i As Long
    Dim userForm As Object
    Dim listData As Object
    Dim rowData As Variant
    
    ' メインシート取得
    Set wsBilling = newBook.Sheets(1)
    lastRow = wsBilling.Cells(Rows.Count, "D").End(xlUp).Row

    ' Dictionary でリストを管理
    Set listData = CreateObject("Scripting.Dictionary")

    ' 現在の請求月取得
    Dim currentBillingMonth As String
    currentBillingMonth = wsBilling.Cells(2, 2).Value ' GYYMM

    ' 該当調剤月以外のデータをリスト化
    For i = 2 To lastRow
        If wsBilling.Cells(i, 2).Value <> currentBillingMonth Then
            rowData = Array(wsBilling.Cells(i, 2).Value, wsBilling.Cells(i, 4).Value, wsBilling.Cells(i, 5).Value, wsBilling.Cells(i, 6).Value)
            listData.Add i, rowData
        End If
    Next i

    ' リストにデータがあればフォーム表示
    If listData.Count > 0 Then
        Set userForm = CreateRebillSelectionForm(listData)
        userForm.Show
    Else
        MsgBox "該当するデータはありません。", vbInformation, "確認"
    End If
End Sub