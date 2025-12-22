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
