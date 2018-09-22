'ダイアログで取込たいファイルを指定する
Sub impt_from_file()
    Dim ask As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "読み込みたいファイルを指定してください"
        With .Filters
            .Clear
            .Add "ファイル", "*.拡張子"
        End With
        .InitialFileName = デフォルトパス
        .AllowMultiSelect = False
        .ButtonName = "読み込む"
        If .Show = True Then
        {ファイル取込操作} .SelectedItems(1)
        End If
    End With
End Sub