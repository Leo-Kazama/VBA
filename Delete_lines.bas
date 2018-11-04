Attribute VB_Name = "Module1"
'いらない行を削除するマクロ
Sub Delete_lines()
    Dim del_line_end As Long
    del_line_end = Range("A1:G300").Find("必要な行").Row - 1
    Range("1:" & del_line_end).Delete
End Sub
