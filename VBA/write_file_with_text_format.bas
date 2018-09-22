'テキストをファイルへ書き込む
Sub txt_writer(content As String, file_name As String, chr As String, ls As Variant, Optional is_wo_bom As Boolean = False)
    '改行コード入りのテキストデータをファイルに書き込む
    Dim strm As ADODB.Stream, buf As Variant
        Set strm = CreateObject("ADODB.Stream")
        With strm
            .Type = adTypeText
            .Charset = chr
            .LineSeparator = ls
            .Open
            .WriteText content
            'ファイル内容を3バイトずらす
            If is_wo_bom = True Then
                .Position = 0
                .Type = adTypeBinary
                .Position = 3
                buf = .read
                .Position = 0
                .write buf
            End If
            .SetEOS
            .SaveToFile file_name, adSaveCreateOverWrite
            .Close
        End With
        Set strm = Nothing
End Sub
