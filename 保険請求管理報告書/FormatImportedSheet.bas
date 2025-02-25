Sub FormatImportedSheet(ws As Worksheet, fileType As String)
    With ws.Cells
        .Font.Name = "MS UI Gothic"
        .Font.Size = 10
        .Columns.AutoFit
    End With

    ' CSVの種類に応じた書式設定（例）
    Select Case fileType
        Case "振込額明細書"
            ws.Rows("1:1").Font.Bold = True
        Case "増減点連絡書"
            ws.Rows("1:1").Interior.Color = RGB(220, 230, 241) ' 青系背景
        Case "返戻内訳書"
            ws.Rows("1:1").Interior.Color = RGB(255, 230, 204) ' オレンジ背景
    End Select
End Sub