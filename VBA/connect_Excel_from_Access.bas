'AccessのデータをExcelから検索して呼び出す
Function find_itm(rng As Range)
    Dim rs As DAO.Recordset, buf As String, acdb As DAO.Database
    buf = Trim(rng.Value)
    If buf <> "" Then
        Set acdb = DBEngine.Workspaces(0).OpenDatabase(“\\共有.accdb")
        Set rs = acdb.OpenRecordset(“Select 知りたい情報 From テーブル Where buf", dbOpenSnapshot)
        If rs.EOF = True And rs.BOF = True Then
            find_itm = "N/A"
        Else
            find_itm = rs!知りたい情報  
        End If
        rs.Close
        Set rs = Nothing
        Set acdb = Nothing
    Else
        find_itm = ""
    End If
    
End Function