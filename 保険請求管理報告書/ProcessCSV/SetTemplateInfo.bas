Sub SetTemplateInfo(newBook As Workbook, targetYear As String, targetMonth As String)
    Dim wsTemplate As Worksheet, wsTemplate2 As Worksheet
    Dim receiptYear As Integer, receiptMonth As Integer
    Dim sendMonth As Integer, sendDate As String

    ' **西暦年と調剤月の計算**
    receiptYear = CInt(targetYear)
    receiptMonth = CInt(targetMonth)

    ' **請求月の計算**
    sendMonth = receiptMonth + 1
    If sendMonth = 13 Then sendMonth = 1
    sendDate = sendMonth & "月10日請求分"

    ' **シートA, Bを取得**
    Set wsTemplate = newBook.Sheets("A")
    Set wsTemplate2 = newBook.Sheets("B")

    ' **シート名変更**
    wsTemplate.Name = "R" & (receiptYear - 2018) & "." & receiptMonth
    wsTemplate2.Name = ConvertToCircledNumber(receiptMonth)

    ' **情報転記**
    wsTemplate.Range("G2").Value = targetYear & "年" & targetMonth & "月調剤分"
    wsTemplate.Range("I2").Value = sendDate
    wsTemplate.Range("J2").Value = ThisWorkbook.Sheets(1).Range("B1").Value
    wsTemplate2.Range("H1").Value = targetYear & "年" & targetMonth & "月調剤分"
    wsTemplate2.Range("J1").Value = sendDate
    wsTemplate2.Range("L1").Value = ThisWorkbook.Sheets(1).Range("B1").Value
End Sub