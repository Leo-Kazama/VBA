Attribute VB_Name = "Module1"
'いらない行を削除してリンクを作成
Sub Delete_lines_add_link()
    Call Delete_lines
    Range("F1").Value = "=HYPERLINK(""..\html\""&$B$2&"".html"",""リンク"")"
End Sub

'いらない行を削除
Sub Delete_lines()
    Dim del_line_end As Long
    del_line_end = Range("A1:G300").Find("必要な行").Row - 1
    Range("1:" & del_line_end).Delete
End Sub
