Sub �f�[�^����()
'
' �f�[�^���� Macro
' �I�������͈͂̃f�[�^��10000�s�ɑ������܂��B1�s�ڂ͗񖼁A2�s�ڂ���f�[�^���n�܂�z��ł��B
'
    start_column = Selection(1).Column
    end_column = Selection(Selection.Count).Column
    end_row = Selection(Selection.Count).Row
    Range(Cells(2, start_column), Cells(end_row, end_column)).Select '�R�s�[���͈̔͂�I��
    '�I�[�g�t�B����K�p����͈͂��w��
    Selection.AutoFill Destination:=Range(Cells(2, start_column), Cells(10001, end_column)), Type:=xlFillDefault
End Sub
