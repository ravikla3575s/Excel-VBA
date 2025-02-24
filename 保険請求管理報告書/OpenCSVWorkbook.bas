Function OpenCSVWorkbook(csvFilePath As String) As Workbook
    On Error Resume Next
    Set OpenCSVWorkbook = Workbooks.Open(csvFilePath)
    On Error GoTo 0
End Function