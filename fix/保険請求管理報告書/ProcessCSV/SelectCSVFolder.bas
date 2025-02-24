Function SelectCSVFolder() As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "CSVフォルダを選択してください"
        If .Show = -1 Then
            SelectCSVFolder = .SelectedItems(1) & "\"
        Else
            MsgBox "フォルダが選択されませんでした。処理を中止します。", vbExclamation
            SelectCSVFolder = ""
        End If
    End With
End Function