'リンクテーブルのリンク先を変更する
'リンク先を本番/ローカルパスへ変更
Function chg_link(Optional is_desktop_path As Boolean = False)
    Dim tdf As dao.TableDef, fld_path As String
    If is_desktop_path = False Then
        fld_path = "共有フォルダパス"
    Else
        'デスクトップパス
        fld_path = "デスクトップパス"
    End If
    fld_path = fld_path & "ファイル.accdb"
    'エラートラップ
    If Dir(fld_path) = "" Then
        MsgBox "接続先にファイル.accdbがありません"
        chg_link = "NG"
        Exit Function
    End If
    If InStr(link_info, fld_path) = 0 Then
        For Each tdf In CurrentDb.TableDefs
            If tdf.Connect <> "" Then
                tdf.Connect = ";DATABASE=" & fld_path & ";TABLE=" & tdf.Name
                tdf.RefreshLink
            End If
        Next tdf
        Debug.Print "done: " & fld_path
    End If
    chg_link = "OK"
End Function

'DBリンク先を返す
Function link_info()
    Dim tdf As TableDef
    Set tdf = CurrentDb.TableDefs("リンクテーブル名")
    link_info = tdf.Connect
    Set tdf = Nothing
End Function