Attribute VB_Name = "PythonFromVBA"
' xlwings アドインがインストールされている必要があります
' VScodeターミナルに xlwings addin install 入力するとインストールできます。
' ツールの参照設定でxlwingsにチェックを入れてください。
' 環境変数に変数名: PYTHONPATH  変数値:C:\Users\*****\Pythonを追加してください。

Sub Run_CallModifyChart()
Attribute Run_CallModifyChart.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' Keyboard Shortcut: Ctrl+Q
'
    RunPython ("import Call_ModifyChart; Call_ModifyChart.main()")
End Sub

Sub Python2()
Attribute Python2.VB_ProcData.VB_Invoke_Func = "w\n14"
'
' Keyboard Shortcut: Ctrl+W
'
    RunPython ("import Call_ModifyChart; Call_ModifyChart.main()")
End Sub

Sub Python3()
Attribute Python3.VB_ProcData.VB_Invoke_Func = "e\n14"
'
' Keyboard Shortcut: Ctrl+E
'
    RunPython ("import Call_ModifyChart; Call_ModifyChart.main()")
End Sub

Sub 選択されているセル範囲内の図形を削除する()
Attribute 選択されているセル範囲内の図形を削除する.VB_ProcData.VB_Invoke_Func = "m\n14"
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
