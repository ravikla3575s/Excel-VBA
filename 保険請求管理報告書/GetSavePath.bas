Function GetSavePath() As String
    GetSavePath = ThisWorkbook.Sheets(1).Range("B3").Value
End Function