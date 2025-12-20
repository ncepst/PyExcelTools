#ScatterChart.py
import xlwings as xw
from xlwings.constants import AxisType
from win32com.client import constants

def ScatterChart(ws,
                 start_range,    # "H3"など
                 row,
                 col,
                 paste_range,    # "A1"など
                 width_cm=12.54,
                 height_cm=7.54,
                 name = "グラフ名",
                 Title = "",
                 SeriesName = "",
                 RGBcolor=(68, 114, 196),
                 x_title = "",
                 x_min =-90,
                 x_max =+90,
                 x_major=15,
                 x_cross=-90,
                 x_format="0",
                 y_title = "",
                 y_min ="",
                 y_max ="",
                 y_major="",
                 y_cross="",
                 y_format="0.0",
                 legend=""):
    
    # RGBのヘルパー関数
    def RGB(r, g, b):
        return r + g*256 + b*65536

    # cm → pt 換算関数の定義 (1 point = 1/72 inch, 1 inch = 2.54 cm)
    def cm_to_pt(cm):
        return cm * 72 / 2.54
    
    color = RGB(*RGBcolor)
        
    # ----------------------------------------------------------
    # 散布図のエクセルグラフを作成する
    # ----------------------------------------------------------
    # (セル範囲入力) --------------------------------------------
    start_range = start_range
    start = ws[start_range]
    row = row
    col = col
    # (ターゲットセル計算) --------------------------------------
    target_row = start.row + row - 1
    target_col = xw.utils.col_name(start.column + col - 1)
    target_range = f"{target_col}{target_row}"

    # xlwings によるグラフ作成 ----------------------------------
    paste_range = paste_range
    chart = ws.charts.add(left=ws.range(paste_range).left+1,  # leftとtopは貼り付け位置の指定 (必須)
                        top=ws.range(paste_range).top+1,
                        width=cm_to_pt(width_cm),       # widthとheightは大きさ指定 (省略可)
                        height=cm_to_pt(height_cm)) 
    chart.chart_type = 'xy_scatter_lines'
    chart.set_source_data(ws.range(f'{start_range}:{target_range}'))
    # ----------------------------------------------------------

    # グラフのチャート名 (エクセル画面左上の表示で確認できる)
    chart.name = name

    # ChartObject(枠) → api[0]、Chart本体 → api[1]
    ch = chart.api[1]

    # グラフタイトル
    if Title =="":
        ch.HasTitle = True
        ch.ChartTitle.Text = ""
    else:
        ch.HasTitle = True
        ch.ChartTitle.Text = Title

    # 横軸のオプション
    x_axis = ch.Axes(AxisType.xlCategory)
    if x_min == "":     #最小値
        x_axis.MinimumScaleIsAuto = True
    else:
        x_axis.MinimumScale = x_min
    if x_max == "":     #最大値
        x_axis.MaximumScaleIsAuto = True
    else:
        x_axis.MaximumScale = x_max
    if x_major !="":     # 目盛間隔
        x_axis.MajorUnit = x_major    
    if x_cross !="":     # 交差位置(縦軸との交点)
        x_axis.CrossesAt = x_cross
    if x_format !="": 
        x_axis.TickLabels.NumberFormatLocal = x_format

    # 横軸のタイトル
    if x_title =="":
        x_axis.HasTitle = False
    else:
        x_axis.HasTitle = True
        x_axis.AxisTitle.Text = x_title

    # 縦軸のオプション
    y_axis = ch.Axes(AxisType.xlValue)
    if y_min == "":     #最小値
        y_axis.MinimumScaleIsAuto = True
    else:
        y_axis.MinimumScale = y_min
    if y_max == "":     #最大値
        y_axis.MaximumScaleIsAuto = True
    else:
        y_axis.MaximumScale = y_max
    if y_major !="":     # 目盛間隔
        y_axis.MajorUnit = y_major    
    if y_cross !="":     # 交差位置(縦軸との交点)
        y_axis.CrossesAt = y_cross
    if y_format !="": 
        y_axis.TickLabels.NumberFormatLocal = y_format

    # 横軸のタイトル
    if y_title =="":
        y_axis.HasTitle = False
    else:
        y_axis.HasTitle = True
        y_axis.AxisTitle.Text = y_title

    # Excel 2021 以降の標準スタイルを指定する ----------------------------------
    # グラフタイトルの文字色をRGB(89,89,89)とし、フォントサイズを14にする
    if ch.HasTitle:
        ch.ChartTitle.Format.TextFrame2.TextRange.Font.Bold = False
        ch.ChartTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(89,89,89)
        ch.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 14

    # グリッド線の設定（薄いグレー)
    x_axis.HasMajorGridlines = True
    x_axis.MajorGridlines.Format.Line.ForeColor.RGB = RGB(217, 217, 217)
    x_axis.MajorGridlines.Format.Line.Weight = 0.75
    x_axis.Format.Line.ForeColor.RGB = RGB(191, 191, 191)
    x_axis.MajorTickMark = constants.xlTickMarkNone # 目盛の内向き/外向きなし
    y_axis.HasMajorGridlines = True
    y_axis.MajorGridlines.Format.Line.ForeColor.RGB = RGB(217, 217, 217)
    y_axis.MajorGridlines.Format.Line.Weight = 0.75
    y_axis.Format.Line.ForeColor.RGB = RGB(191, 191, 191)
    y_axis.MajorTickMark = constants.xlTickMarkNone # 目盛の内向き/外向きなし

    for ax in ch.Axes():    
        # 軸の設定
        ax.TickLabels.Font.Color = RGB(89,89,89)
        ax.TickLabels.Font.Size = 9
        ax.TickLabels.Font.Name = "Aptos Narrow 本文"
        # 軸タイトルがあるとき、軸タイトルを設定する
        if ax.HasTitle:
            ax.AxisTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(89,89,89)
            ax.AxisTitle.Format.TextFrame2.TextRange.Font.Bold = False
            ax.AxisTitle.Format.TextFrame2.TextRange.Font.Size = 10

    # 外枠の設定
    ch.ChartArea.Format.Line.ForeColor.RGB = RGB(217,217,217) # 薄いグレー
    ch.ChartArea.Format.Line.Weight = 0.75                    # 枠線の太さ(pt)
    # ----------------------------------------------------------------------

    # 凡例なし
    if legend=="":
        ch.HasLegend = False
    else:
        ch.HasLegend = True
        if legend=="auto":
            pass
        if "U" in legend:
            ch.Legend.Top = ch.PlotArea.InsideTop
        if "R" in legend: 
            ch.Legend.Left = ch.PlotArea.InsideLeft + ch.PlotArea.InsideWidth - ch.Legend.Width
        if "10" in legend:
            ch.Legend.Format.TextFrame2.TextRange.Font.Size = 10          
        
    # プロットエリアの調整
    p = ch.PlotArea
    p.InsideLeft   = p.InsideLeft
    p.InsideTop    = p.InsideTop
    p.InsideWidth  = p.InsideWidth
    p.InsideHeight = p.InsideHeight+15  #下側に広げる

    # 1つ目の系列の色を青にする
    series = ch.SeriesCollection(1)
    if not SeriesName == "":
        series.Name = SeriesName
    series.Format.Line.ForeColor.RGB = color             # 線の色
    series.MarkerForegroundColor = color                 # マーカー枠線の色
    series.MarkerBackgroundColor = color                 # マーカー内部の色

    # Excel 2021 以降の標準スタイルを指定する ---------------------------------
    # 線とマーカーの設定
    series.Format.Line.Weight = 1.5                     # 線の太さ(pt)
    series.MarkerStyle = constants.xlMarkerStyleCircle  # マーカー: 丸
    series.MarkerSize = 5                               # マーカーサイズ
    # -----------------------------------------------------------------------

    # 軸タイトルと目盛りの数値の色を黒に変更する
    for ax in ch.Axes():
        ax.TickLabels.Font.Color = RGB(0,0,0)
        if ax.HasTitle:
            ax.AxisTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0,0,0)
            
    # グラフ外枠を黒に変更
    ch.ChartArea.Format.Line.ForeColor.RGB = RGB(0,0,0)
    
    return ch