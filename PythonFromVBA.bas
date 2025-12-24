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
    RunPython ("import Call_ModifyChart2; Call_ModifyChart2.main()")
End Sub

Sub Python3()
'
' Keyboard Shortcut: Ctrl+E
'
    RunPython ("import Call_ModifyChart3; Call_ModifyChart3.main()")
End Sub

' 便利マクロ
Sub グラフ作成_自動判定()

' Keyboard Shortcut: Ctrl＋G
'
    Dim ws As Worksheet
    Dim rngStart As Range
    Dim rng As Range
    Dim lastRow As Long, lastCol As Long
    Dim chartObj As ChartObject
    Dim i As Long
    Dim rowCount As Long, colCount As Long
    Dim seriesByRow As Boolean
    
    Set ws = ActiveSheet
    Set rngStart = Selection.Cells(1, 1)

    ' 連続データ範囲を自動判定
    lastRow = rngStart.End(xlDown).Row
    lastCol = rngStart.End(xlToRight).Column
    Set rng = ws.Range(rngStart, ws.Cells(lastRow, lastCol))

    rowCount = rng.Rows.count
    colCount = rng.Columns.count

    ' ---- 系列方向の自動判定 ----
    ' 横が長ければ「系列は横」
    seriesByRow = (colCount > rowCount)

    ' グラフ作成
    Set chartObj = ws.ChartObjects.Add(Left:=100, Top:=50, Width:=400, Height:=300)

    With chartObj.Chart
        .ChartType = xlXYScatterLines

        ' 既存系列削除
        Do While .SeriesCollection.count > 0
            .SeriesCollection(1).Delete
        Loop

        If Not seriesByRow Then
            ' ===== 系列：縦（通常）=====
            ' 1列目：X、2列目以降：Y
            For i = rngStart.Column + 1 To lastCol
                With .SeriesCollection.NewSeries
                    .Name = ws.Cells(rngStart.Row, i).Value
                    .XValues = ws.Range(ws.Cells(rngStart.Row + 1, rngStart.Column), _
                                        ws.Cells(lastRow, rngStart.Column))
                    .Values = ws.Range(ws.Cells(rngStart.Row + 1, i), _
                                       ws.Cells(lastRow, i))
                End With
            Next i

        Else
            ' ===== 系列：横 =====
            ' 1行目：X、2行目以降：Y
            For i = rngStart.Row + 1 To lastRow
                With .SeriesCollection.NewSeries
                    .Name = ws.Cells(i, rngStart.Column).Value
                    .XValues = ws.Range(ws.Cells(rngStart.Row, rngStart.Column + 1), _
                                        ws.Cells(rngStart.Row, lastCol))
                    .Values = ws.Range(ws.Cells(i, rngStart.Column + 1), _
                                       ws.Cells(i, lastCol))
                End With
            Next i
        End If

        .HasTitle = True
        .ChartTitle.Text = "グラフ タイトル"
        .HasLegend = False
        
        '.Axes(xlCategory).HasTitle = True
        '.Axes(xlCategory).AxisTitle.Text = "X軸"
        '.Axes(xlValue).HasTitle = True
        '.Axes(xlValue).AxisTitle.Text = "Y軸"
    End With
End Sub

Sub グラフ作成_選択範囲()
'
' Keyboard Shortcut: Ctrl+Shift＋G
'
    Dim ws As Worksheet
    Dim rng As Range
    Dim chartObj As ChartObject
    Dim i As Long
    Dim rowCount As Long, colCount As Long
    Dim seriesByRow As Boolean

    Set ws = ActiveSheet

    If TypeName(Selection) <> "Range" Then Exit Sub
    Set rng = Selection

    rowCount = rng.Rows.count
    colCount = rng.Columns.count

    ' 横が長ければ「系列は横」
    seriesByRow = (colCount > rowCount)

    Set chartObj = ws.ChartObjects.Add(Left:=100, Top:=50, Width:=400, Height:=300)

    With chartObj.Chart
        .ChartType = xlXYScatterLines

        ' 既存系列削除
        Do While .SeriesCollection.count > 0
            .SeriesCollection(1).Delete
        Loop

        If Not seriesByRow Then
            ' ===== 系列：縦 =====
            For i = 2 To colCount
                With .SeriesCollection.NewSeries
                    .Name = rng.Cells(1, i).Value
                    .XValues = rng.Columns(1).Offset(1).Resize(rowCount - 1)
                    .Values = rng.Columns(i).Offset(1).Resize(rowCount - 1)
                End With
            Next i
        Else
            ' ===== 系列：横 =====
            For i = 2 To rowCount
                With .SeriesCollection.NewSeries
                    .Name = rng.Cells(i, 1).Value
                    .XValues = rng.Rows(1).Offset(, 1).Resize(, colCount - 1)
                    .Values = rng.Rows(i).Offset(, 1).Resize(, colCount - 1)
                End With
            Next i
        End If

        .HasTitle = True
        .ChartTitle.Text = "グラフ タイトル"
        .HasLegend = False
    End With
End Sub

Sub 選択されているセル範囲内の図形を削除()
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

Sub  選択されているセル範囲内の図形をグループ化()
'
' Keyboard Shortcut: Ctrl＋U
'
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim selRange As Range
    Set selRange = Selection
    
    Dim selLeft As Double, selTop As Double, selRight As Double, selBottom As Double
    selLeft = selRange.Left
    selTop = selRange.Top
    selRight = selRange.Left + selRange.Width
    selBottom = selRange.Top + selRange.Height
    
    Dim shp As Shape
    Dim shpNames() As String
    Dim count As Long
    count = 0
    
    ' 選択セルに重なる図形を配列に追加
    For Each shp In ws.Shapes
        Dim shpLeft As Double, shpTop As Double, shpRight As Double, shpBottom As Double
        shpLeft = shp.Left
        shpTop = shp.Top
        shpRight = shp.Left + shp.Width
        shpBottom = shp.Top + shp.Height
        
        If Not (shpRight < selLeft Or shpLeft > selRight Or shpBottom < selTop Or shpTop > selBottom) Then
            count = count + 1
            ReDim Preserve shpNames(1 To count)
            shpNames(count) = shp.Name
        End If
    Next shp
    
    ' 2つ以上あればグループ化
    If count >= 2 Then
        ' グループ化は.Group  選択は.Select
        ws.Shapes.Range(shpNames).Group
    End If
End Sub

Sub 選択範囲の数式を値として貼り付け()
'
' Keyboard Shortcut: Ctrl+Shift+M
'
    Selection.Value = Selection.Value
End Sub

Sub 表示する小数桁の設定()
'
' Keyboard Shortcut:
'
    Selection.NumberFormat = "0.000"
End Sub


