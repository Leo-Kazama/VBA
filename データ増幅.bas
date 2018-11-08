Sub データ増幅()
'
' データ増幅 Macro
' 選択した範囲のデータを10000行に増幅します。1行目は列名、2行目からデータが始まる想定です。
'
    start_column = Selection(1).Column
    end_column = Selection(Selection.Count).Column
    end_row = Selection(Selection.Count).Row
    Range(Cells(2, start_column), Cells(end_row, end_column)).Select 'コピー元の範囲を選択
    'オートフィルを適用する範囲を指定
    Selection.AutoFill Destination:=Range(Cells(2, start_column), Cells(10001, end_column)), Type:=xlFillDefault
End Sub
