Function GetTemplatePath() As String
    GetTemplatePath = ThisWorkbook.Sheets(1).Range("B2").Value & "\保険請求管理報告書テンプレート20250222.xltm"
End Function