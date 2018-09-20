'クリップボードのデータをアプリとOS間でやり取りする
'クリップボード制御　引数の有無で引き出す/格納
Function clip_IO(Optional input_str As Variant)
    Dim Dobj As DataObject
    Set Dobj = New DataObject
    If IsMissing(input_str) Then
        Dobj.GetFromClipboard
        clip_IO = Dobj.GetText
    Else
        Dobj.SetText input_str
        Dobj.PutInClipboard
    End If
    Set Dobj = Nothing
End Function