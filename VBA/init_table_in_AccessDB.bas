'指定テーブルを確認して削除する
'=>クエリ処理行う際に一時テーブルの初期化に利用

'指定テーブルを初期化 ※削除
Sub tbl_init(tbl As String)
    Dim tbl_arr() As String, i As Integer
    tbl_arr = Split(tbl, ",")
    For i = 0 To UBound(tbl_arr)
        If sch_tbl(tbl_arr(i)) = True Then
            'テーブルが開いている場合テーブルを閉じる
            If SysCmd(acSysCmdGetObjectState, acTable, tbl_arr(i)) = 1 Then
                DoCmd.Close acTable, tbl_arr(i)
            End If
            DoCmd.DeleteObject acTable, tbl_arr(i)
        End If
    Next i
End Sub

'テーブルの有無を確認
Function sch_tbl(tbl_name As String)
    Dim t_def As DAO.TableDef, tmp As Boolean
    tmp = False
    For Each t_def In CurrentDb.TableDefs
        If t_def.Name Like tbl_name Then
            tmp = True
        End If
    Next
    sch_tbl = tmp
End Function
