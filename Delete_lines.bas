Attribute VB_Name = "Module1"
'����Ȃ��s���폜����}�N��
Sub Delete_lines()
    Dim del_line_end As Long
    del_line_end = Range("A1:G300").Find("�K�v�ȍs").Row - 1
    Range("1:" & del_line_end).Delete
End Sub
