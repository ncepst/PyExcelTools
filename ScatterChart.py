#ScatterChart.py
import xlwings as xw
from xlwings.constants import AxisType
from win32com.client import constants
import re

# ScatterChart : 新規のグラフ作成
# from ScatterChart import ScatterChart

def ScatterChart(ws,
                 start_range = "A1",
                 row = 2,
                 col = 2,
                 paste_range = "A1",
                 width_cm = 12.54,
                 height_cm = 7.54,
                 name = "",
                 Title = "",
                 series_list = None,
                 x_title = "",
                 x_min = "",
                 x_max = "",
                 x_major = "",
                 x_cross = "",
                 x_format = "",
                 y_title = "",
                 y_min = "",
                 y_max = "",
                 y_major = "",
                 y_cross = "",
                 y_format = "",
                 legend = "",
                 chart_border_color = None,  #None=dafault, 黒枠=0, 枠なし=False
                 Title_space   = +0,
                 x_title_space = -15, #下に広げる場合はマイナス
                 y_title_space = +0,
                 Title_color = "",
                 Title_size = "",
                ):
    
    # list / dict はミュータブルのため、デフォルト引数を None にしている
    if series_list is None:
        series_list = [{"color_RGB": (68,114,196)}]
    """
    注意) series_listの系列数をデータ範囲以下にしないと例外発生となる
    series_list = [{"name":"系列1", "color_RGB": (68,114,196)},  # 青
                   {"name":"系列2", "color_RGB": (237,125,49)},  # オレンジ
                   {"name":"系列3", "color_RGB": (112,173,71)},  # 緑
                   {"name":"系列4", "color_RGB": (165,165,165)}, # グレー
                  ],
    """                
    # RGBのヘルパー関数
    def RGB(r, g, b):
        return r + g*256 + b*65536
    
    # cm → pt 換算関数の定義 (1 point = 1/72 inch, 1 inch = 2.54 cm)
    def cm_to_pt(cm):
        return cm * 72 / 2.54
        
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
    if name!="":
        chart.name = name

    # ChartObject(枠) → api[0]、Chart本体 → api[1]
    ch = chart.api[1]

    # グラフタイトル
    if Title =="" or Title == False:
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
    if x_title == "" or x_title == False:
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
    if y_title == "" or y_title == False:
        y_axis.HasTitle = False
    else:
        y_axis.HasTitle = True
        y_axis.AxisTitle.Text = y_title
    
    # 系列の設定 -----------------------------------------------------------------------------
    marker_map = {
            "C":constants.xlMarkerStyleCircle,
            "S":constants.xlMarkerStyleSquare,
            "D":constants.xlMarkerStyleDiamond,
            "T":constants.xlMarkerStyleTriangle
            }
    
    for i, cfg in enumerate(series_list, start=1):
        try:
            series = ch.SeriesCollection(i)
            name0 = cfg.get("name")
            if name0 not in (None, ""):
                series.Name = cfg["name"]
            if cfg.get("XValues"):
                series.XValues = ws.range(cfg["XValues"] ).api
            if cfg.get("Values"):
                series.Values  = ws.range(cfg["Values"]).api
            color = cfg.get("color_RGB")
            if color not in (None, ""):
                color_rgb = RGB(*color)
                series.Format.Line.ForeColor.RGB = color_rgb    # 線の色
                series.MarkerForegroundColor = color_rgb        # マーカー枠線の色
                series.MarkerBackgroundColor = color_rgb        # マーカー内部の色
            
            smooth = cfg.get("smooth", True)
            series.Smooth = bool(smooth)
            
            if cfg.get("marker"):
                cfg["marker"] = marker_map.get(cfg["marker"], marker_map["C"])
            
            # デフォルト値は Excel 2021 以降の標準スタイル
            style = cfg.get("style") or "line+marker"
            if "marker" in style:
                series.MarkerStyle = cfg.get("marker",marker_map["C"])  # マーカー: 丸
                series.MarkerSize = cfg.get("size",5)                   # マーカーサイズ
            else:
                series.MarkerStyle = constants.xlMarkerStyleNone
            if "line" in style:
                series.Format.Line.Visible = True
                series.Format.Line.Weight = cfg.get("weight", 1.5)  # 線の太さ(pt)
            elif style.startswith("dash"):
                series.Format.Line.Visible = True
                series.Format.Line.DashStyle = 4
                series.Format.Line.Weight = cfg.get("weight", 1.5)
            elif style.startswith("chain"):
                series.Format.Line.Visible = True
                series.Format.Line.DashStyle = 5
                series.Format.Line.Weight = cfg.get("weight", 1.5)
            else:
                series.Format.Line.Visible = False
                
            alpha = cfg.get("alpha") # 透明度は0~1
            if alpha not in (None, ""):
                series.Format.Line.Transparency = float(alpha)
                
        except Exception as e:
            print(f"系列{i}で例外発生:{e}")
    #-----------------------------------------------------------------------------------------
    
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
        # "Aptos Narrow 本文"は Excel 2021以降のみ
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

    # 軸タイトルと目盛りの数値の色を黒に変更する
    for ax in ch.Axes():
        ax.TickLabels.Font.Color = RGB(0,0,0)
        if ax.HasTitle:
            ax.AxisTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0,0,0)
            
    # グラフ外枠の色を変更
    if chart_border_color not in (None, ""):
        if isinstance(chart_border_color, (tuple, list)):
            chart_border_color = RGB(*chart_border_color)
        ch.ChartArea.Format.Line.ForeColor.RGB = chart_border_color
    elif chart_border_color == False:
        # グラフの外枠を消す
        ch.ChartArea.Border.LineStyle = 0
    
    # タイトルの色を変更
    if Title_color not in (None, ""):
        if isinstance(Title_color, (tuple, list)):
            Title_color = RGB(*Title_color)
        ch.ChartArea.Format.Line.ForeColor.RGB = Title_color
        
    # タイトルのサイズを変更
    if Title_size != "" and Title_size != 0:
        ch.ChartTitle.Format.TextFrame2.TextRange.Font.Size = Title_size
        
    # 凡例を一度無効にする(例外あり)
    if legend == "right":
        ch.HasLegend = True
    else:
        ch.HasLegend = False
    
    # プロットエリアの調整
    p = ch.PlotArea
    p.InsideLeft   = p.InsideLeft + y_title_space
    p.InsideTop    = p.InsideTop + Title_space
    p.InsideWidth  = p.InsideWidth - y_title_space
    p.InsideHeight = p.InsideHeight - Title_space - x_title_space
    
    # 凡例設定
    if legend == "" or legend == False:
        ch.HasLegend = False
    else:
        ch.HasLegend = True
        if legend == "auto":
            pass
        else:
            if "U" in legend: # 大文字のUを含む場合
                ch.Legend.Top = ch.PlotArea.InsideTop
            if "R" in legend: # 大文字のRを含む場合
                ch.Legend.Left = ch.PlotArea.InsideLeft + ch.PlotArea.InsideWidth - ch.Legend.Width
            m = re.search(r"\d+(?:\.\d+)?", legend)
            if m:
                ch.Legend.Format.TextFrame2.TextRange.Font.Size = float(m.group())
            if "FW" in legend:
                ch.Legend.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
                ch.Legend.Format.Fill.Visible = True
            if "BB" in legend:
                ch.Legend.Format.Line.ForeColor.RGB = RGB(0, 0, 0)
                ch.Legend.Format.Line.Weight = 0.75
                ch.Legend.Format.Line.Visible = True
            if "TB" in legend:
                ch.Legend.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
            
    return chart