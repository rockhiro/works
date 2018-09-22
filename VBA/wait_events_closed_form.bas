'テーブルが閉じされるイベントを待つ
Sub wait_close_form(form_name As String)
    Do
        DoEvents
    Loop Until SysCmd(acSysCmdGetObjectState, acForm, form_name) = 0
End Sub
