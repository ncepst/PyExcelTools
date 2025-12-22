Attribute VB_Name = "PythonFromVBA"
' xlwings アドインがインストールされている必要があります
' VScodeターミナルに xlwings addin install 入力するとインストールできます。
' ツールの参照設定でxlwingsにチェックを入れてください。

Sub Run_CallModifyChart()
    RunPython _
    "import sys;" & _
    "sys.path.append(r'C:\Users\*****\Python');" & _
    "import Call_ModifyChart;" & _
    "Call_ModifyChart.main(mode=3)"
End Sub
