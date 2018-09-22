'画像取得先をHTMLタグで返す (web/共有フォルダ)
Private Function img_src(path As String, is_web As Boolean)
    Dim tmp As String
    tmp = path
    If is_web = False Then
        tmp = "file:///" & Replace(str_enc(path), "%5C", "/")
    End If
    img_src = "<img src='" & tmp & "' alt = '画像を取得できません' style='max-width: 100%;'>"
End Function