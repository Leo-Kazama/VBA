Attribute VB_Name = "Module1"
'����Ȃ��s���폜���ă����N���쐬
Sub Delete_lines_add_link()
    Call Delete_lines
    Range("F1").Value = "=HYPERLINK(""..\html\""&$B$2&"".html"",""�����N"")"
End Sub

'����Ȃ��s���폜
Sub Delete_lines()
    Dim del_line_end As Long
    del_line_end = Range("A1:G300").Find("�K�v�ȍs").Row - 1
    Range("1:" & del_line_end).Delete
End Sub
