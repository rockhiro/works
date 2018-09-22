'エクセルから指定したURLをFireFoxで開く
Sub open_url(rng As Range)
    Dim wsh As Object, buf As String
    Set wsh = CreateObject("Wscript.shell")
    buf = rng.value
    If buf = "null" Then
        MsgBox "このURLは不正です"
        End
    End If
    wsh.Run "firefox -url " & buf, 1
End Sub