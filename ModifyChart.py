#ModifyChart.py
import xlwings as xw
from xlwings.constants import AxisType
from win32com.client import constants
import re

# MofifyChart : 既存グラフの変更
# from ModifyChart import ModifyChart

def ModifyChart( chart,
                 ws ="",
                 width_cm = "",
                 height_cm = "",
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
    
    # グラフ全体のサイズ変更
    if width_cm not in ("", 0, None):
        chart.width = cm_to_pt(width_cm)
    if height_cm not in ("", 0, None):
        chart.height = cm_to_pt(height_cm)  

    # グラフのチャート名 (エクセル画面左上の表示で確認できる)
    if name!="":
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
                try:
                    series.XValues = ws.range(cfg["XValues"] ).api
                except:
                    print("wsの定義が必要です")
            if cfg.get("Values"):
                try:
                    series.Values  = ws.range(cfg["Values"]).api
                except:
                    print("wsの定義が必要です")
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
        
    # 凡例を一度無効にする(例外あり)
    if legend == "right":
        ch.HasLegend = True
    else:
        ch.HasLegend = False
    
    # プロットエリアの調整
    p = ch.PlotArea
    p.InsideLeft   = p.InsideLeft
    p.InsideTop    = p.InsideTop
    p.InsideWidth  = p.InsideWidth
    p.InsideHeight = p.InsideHeight + 15  #下側に広げる
    
    # 凡例設定
    if legend=="":
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