Attribute VB_Name = "PythonFromVBA"
' 標準モジュールにインポートします
' Python実行には xlwings アドインがインストールされている必要があります
' VScodeターミナルに xlwings addin install を入力するとインストールできます
' ツールの参照設定でxlwingsにチェックを入れてください
' 環境変数に変数名: PYTHONPATH  変数値:C:\Users\*****\Pythonを追加してください
' マクロのオプションから、ショートカットキーが設定できます

Sub Run_CallModifyChart()
'
' Keyboard Shortcut: Ctrl+Q
'
    RunPython ("import Call_ModifyChart; Call_ModifyChart.main()")
End Sub

Sub Python2()
'
' Keyboard Shortcut: Ctrl+W
'
    RunPython ("import Call_ModifyChart; Call_ModifyChart.main()")
End Sub

Sub Python3()
'
' Keyboard Shortcut: Ctrl+E
'
    RunPython ("import Call_ModifyChart; Call_ModifyChart.main()")
End Sub

' 便利マクロ6選
Sub グラフ作成()
'
' Keyboard Shortcut: Ctrl+G
'
    Dim ws As Worksheet
    Dim rng As Range
    Dim lastRow As Long, lastCol As Long
    Dim chartObj As ChartObject
    
    Set ws = ActiveSheet
    
    ' 選択範囲の先頭セルを取得
    Set rng = Selection.Cells(1, 1)
    ' 連続データの範囲を自動判定（右方向と下方向）
    lastRow = rng.End(xlDown).Row
    lastCol = rng.End(xlToRight).Column
    Set rng = ws.Range(rng, ws.Cells(lastRow, lastCol))
    
    ' グラフ作成
    Set chartObj = ws.ChartObjects.Add(Left:=100, Top:=50, Width:=400, Height:=300)
    With chartObj.Chart
        .ChartType = xlXYScatterLines
        .SetSourceData Source:=rng
        .ChartTitle.Text = "グラフ タイトル"
        .HasTitle = True
        '.Axes(xlCategory).HasTitle = True
        '.Axes(xlCategory).AxisTitle.Text = "X軸"
        '.Axes(xlValue).HasTitle = True
        '.Axes(xlValue).AxisTitle.Text = "Y軸"
        .HasLegend = False
    End With
End Sub

Sub 選択されているセル範囲内の図形を削除する()
'
' Keyboard Shortcut: Ctrl+M
'
    Dim shp As Shape
    Dim rng As Range
    
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    For Each shp In ActiveSheet.Shapes
    '図形の配置されているセル範囲をオブジェクト変数にセット
    Set rng = Range(shp.TopLeftCell, shp.BottomRightCell)
    '図形の配置されているセル範囲と
    '選択されているセル範囲が重なっているときに図形を削除
    If Not (Intersect(rng, Selection) Is Nothing) Then
    shp.Delete
    End If
    Next
End Sub

Sub 表示小数桁の変更()
'
' Keyboard Shortcut: Ctrl+Shift+M
'
    Selection.NumberFormat = "0.000"
End Sub

Sub 選択範囲の値を一括コピーして数式を削除()
'
' Keyboard Shortcut:
'
    Selection.Value = Selection.Value
End Sub

Sub 新規エクセルにアクティブシートをコピー()
'
' Keyboard Shortcut:
'
    ActiveSheet.Copy
End Sub

Sub 高さ幅を自動調整()
    Selection.Columns.AutoFit
    Selection.Rows.AutoFit
End Sub

