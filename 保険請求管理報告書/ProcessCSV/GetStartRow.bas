Function GetStartRow(ws As Worksheet, category As String) As Long
    Dim lastRow As Long
    Dim i As Long
    Dim keyword As String
    
    ' 最終行取得
    lastRow = ws.Cells(Rows.Count, "H").End(xlUp).Row

    ' キーワードに応じた開始行を決定
    Select Case category
        Case "社保返戻再請求"
            keyword = "国家→医本"
        Case "社保月遅れ請求"
            keyword = "⑨返戻分再請求分"
        Case "国保返戻再請求"
            keyword = "⑨返戻分再請求分（医保）"
        Case "国保月遅れ請求"
            keyword = "⑩月遅れ請求分（医保）"
        Case "社保返戻・査定"
            keyword = "⑪月送り分"
        Case "社保未請求扱い"
            keyword = "⑪社保　返戻・査定"
        Case "国保返戻・査定"
            keyword = "⑫社保　未請求扱い"
        Case "国保未請求扱い"
            keyword = "⑬国保　返戻・査定"
        Case "国保未請求扱い(追加)"
            keyword = "⑭国保　未請求扱い"
        Case Else
            keyword = "" ' キーワードが設定されていない場合
    End Select

    ' 該当なしの場合は最終行の次の行に設定
    If keyword = "" Then
        GetStartRow = lastRow + 1
        Exit Function
    End If

    ' H列を検索（2行目から）
    For i = 2 To lastRow
        If Trim(LCase(ws.Cells(i, 8).Value)) = LCase(keyword) Then
            GetStartRow = i + 1
            Exit Function
        End If
    Next i

    ' キーワードが見つからなかった場合は最終行の次の行
    GetStartRow = lastRow + 1
End Function