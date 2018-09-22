'Accessアプリを利用しているユーザーを調べる
Function U_NAME()
    'ユーザー判断　unlock_list_Tに登録がある場合Trueを返す
    Dim WshNetwork As Object, is_nobody As Boolean
    Set WshNetwork = CreateObject("WScript.Network")
    is_nobody = true
    If DLookup("usr_name", "unlock_list_T", "usr_name = '" & WshNetwork.UserName & "'") <> "" Then
        is_nobody = false
    End If
    Set WshNetwork = Nothing
    U_NAME = is_nobody
End Function
