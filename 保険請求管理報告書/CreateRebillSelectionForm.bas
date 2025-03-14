Function CreateRebillSelectionForm(listData As Object) As Object
    Dim uf As Object
    Dim listBox As Object
    Dim btnOK As Object
    Dim i As Long
    Dim rowData As Variant

    ' UserForm を作成
    Set uf = CreateObject("Forms.UserForm")
    uf.Caption = "返戻再請求の選択"
    uf.Width = 400
    uf.Height = 500

    ' ListBox を追加（複数選択可能）
    Set listBox = uf.Controls.Add("Forms.ListBox.1", "listBox", True)
    listBox.Left = 20
    listBox.Top = 20
    listBox.Width = 350
    listBox.Height = 350
    listBox.MultiSelect = 1 ' 複数選択可能

    ' リストデータ追加
    For i = 0 To listData.Count - 1
        rowData = listData.Items()(i)
        listBox.AddItem rowData(0) & " | " & rowData(1) & " | " & rowData(2) & " | " & rowData(3)
    Next i

    ' OKボタンを追加
    Set btnOK = uf.Controls.Add("Forms.CommandButton.1", "btnOK", True)
    btnOK.Caption = "確定"
    btnOK.Left = 150
    btnOK.Top = 400
    btnOK.Width = 100
    btnOK.Height = 30

    ' ボタンが押された時のイベント処理
    btnOK.OnAction = "ProcessRebillSelection"

    ' UserForm を返す
    Set CreateRebillSelectionForm = uf
End Function