'PowerShellからデスクトップ上のzipファイルを解凍する
Sub unzip(path As String, fl_name As String)
    Dim wsh As Object, cmd As String
    Set wsh = CreateObject("Wscript.shell")
    cmd = "Expand-Archive -path " & path & "\" & fl_name & " -DestinationPath " & path & "\ -force"
    wsh.Run "powershell -ExecutionPolicy RemoteSigned -Command " & cmd, 0
    Set wsh = Nothing
End Sub

'実行環境のデスクトップパスを返す
Function desktop_path()
    Dim wsh As Object
    Set wsh = CreateObject("Wscript.shell")
    desktop_path = wsh.specialfolders("desktop")
    Set wsh = Nothing
End Function

'デスクトップ状のzipファイルを順番に読み込んで実行する
Sub unzip_around()
    Dim fso As FileSystemObject, fl As Variant, fld_path As String
    fld_path = desktop_path
    Set fso = New FileSystemObject
    For Each fl In fso.GetFolder(fld_path).Files
        If fl.Name Like "*.zip" Then
            Call unzip(fld_path, fl.Name)
        End If
    Next fl
    Set fso = Nothing
    MsgBox "解凍完了"
End Sub
