# ModifyChart.py
# Copyright (c) 2025 ncepst
# SPDX-License-Identifier: MIT
"""
概要:
既存の Excel グラフの書式設定を簡単に行うモジュールです。

使用法:
from ModifyChart import ModifyChart, RGB
ModifyChart(chart, preset="std", NS=2) のように呼び出して使用
"""
from xlwings.constants import AxisType
from win32com.client import constants
# 型ヒントには Python 3.10 以降の記法を使用しています

# RGB値をExcel用の整数に変換するヘルパー関数
def RGB(r, g, b):
    return r | (g << 8) | (b << 16)

# cm → pt 換算関数 (1 point = 1/72 inch, 1 inch = 2.54 cm)
def cm_to_pt(cm):
    return cm * 72 / 2.54

def emphasize_line(axis,value=0,weight=1):
    axis.CrossesAt = value
    line = axis.Format.Line
    line.Weight = weight
    line.ForeColor.RGB = RGB(0,0,0)

PRESET = {
    # "Aptos Narrow 本文"は Excel 2021以降のみ
    "excel2021": {
        "title_font_size": 14,
        "title_font_color": RGB(89,89,89),
        "title_font_bold": False,
        "title_font_name": "Aptos Narrow 本文",
        "axis_title_font_size": 10,
        "axis_title_font_name": "Aptos Narrow 本文",
        "axis_title_font_color": RGB(89,89,89),
        "axis_title_font_bold": False,
        "axis_tick_font_color": RGB(89, 89, 89),
        "axis_tick_font_size": 9,
        "axis_tick_font_name": "Aptos Narrow 本文",
        "axis_line": True,
        "axis_line_color": RGB(191, 191, 191),
        "axis_line_weight": 0.75,    
        "x_major_grid": True,
        "y_major_grid": True,
        "x_minor_grid": None,
        "y_minor_grid": None,
        "major_grid_color": RGB(217, 217, 217),
        "major_grid_weight": 0.75,
        # TickMark: None, Inside, Outside, Cross
        "x_major_tickmark": constants.xlTickMarkNone,  # 目盛の内向き/外向きなし
        "y_major_tickmark": constants.xlTickMarkNone,
        "x_minor_tickmark": None,
        "y_minor_tickmark": None,
        # 外枠の有無は引数で指定してください
        "frame_color": RGB(217,217,217),
        "frame_weight": 0.75,
        "plot_area_frame": None,
        "plot_area_frame_color": RGB(0,0,0),
        "plot_area_frame_weight": 1.0,
        "style": "line+marker",
        "smooth": True,
        "alpha": None,
        "line_weight": 1.5,
        "marker": "C",
        "marker_size": 5,
        "y2_major_grid": False,  # 副軸グリッド表示なし
        "y2_minor_grid": False,
        "y2_major_tickmark": constants.xlTickMarkNone,
        "y2_minor_tickmark": None,
    },
    "std": {
        "title_font_color": RGB(0,0,0),
        "axis_title_font_color": RGB(0,0,0), 
        "axis_tick_font_color": RGB(0,0,0),
    },
    "no_grid":{
        "x_major_grid": False,
        "y_major_grid": False,
        "plot_area_frame": True,
        "axis_line": True,
        "axis_line_color": RGB(0, 0, 0),  
        "x_major_tickmark": constants.xlTickMarkInside,
        "y_major_tickmark": constants.xlTickMarkInside,
    },
    "grid":{
        "major_grid_color": RGB(0, 0, 0),
        "axis_line_color": RGB(0, 0, 0),
    },
}
# std を excel2021 ベースで上書き
PRESET["std"] = {**PRESET["excel2021"], **PRESET["std"]}
# std ベースで上書き
PRESET["no_grid"] = {**PRESET["std"], **PRESET["no_grid"]}
PRESET["grid"] = {**PRESET["std"], **PRESET["grid"]}

marker_map = {
        "C":constants.xlMarkerStyleCircle,
        "S":constants.xlMarkerStyleSquare,
        "D":constants.xlMarkerStyleDiamond,
        "T":constants.xlMarkerStyleTriangle,
        "N":constants.xlMarkerStyleNone
        }

# False:無効化, None: 変更なし もしくは デフォルト値
# 優先順位 series cfg > 引数 > preset
def ModifyChart(chart,                        # ExcelのChartオブジェクト
                ws = None,                    # None (系列ごとにRange指定する場合のみWorksheetオブジェクトを指定)
                preset = "std",               # プリセットスタイル名
                width_cm = None,              # グラフ幅(cm), Noneで変更なし
                height_cm = None,             # グラフ高さ(cm), Noneで変更なし
                name = None,                  # グラフ名変更, Noneで変更なし
                title: str|bool|None = None,  # タイトル文字列, Falseで無効化, Noneで変更なし
                title_font_color = None,      # タイトルフォント色(RGB)
                title_font_size = None,       # タイトルフォントサイズ
                title_space = +0,             # タイトルとグラフの間隔(pt)
                NS = 1,                       # データ系列数
                series_list: list[dict]|None = None,  # 各系列の設定
                style = None,                 # 線＋マーカーのスタイル
                smooth = None,                # Trueで曲線、Falseで折れ線
                marker = None,                # マーカー種類: "C":●, "S":■, "D":◆, "T":▲, "N":なし
                alpha = None,                 # 線の透明度(0~1)
                x_title: str|bool|None = None,# X軸タイトル文字列。Falseで無効化、Noneで変更なし
                x_title_space = +0,           # プロットエリアを下側に広げる場合はマイナス
                x_min: float|str|None = None, # X軸最小値, "auto"で自動調整, Noneで変更なし
                x_max: float|str|None = None, # X軸最大値, "auto"で自動調整, Noneで変更なし
                x_major: float|None = None,   # X軸主目盛間隔, Noneで変更なし, グリッド線表示設定はPRESET
                x_minor: float|None = None,   # X軸副目盛間隔, Noneで変更なし, グリッド線表示設定はPRESET
                x_cross = None,               # Y軸との交差位置, Noneで変更なし
                x_format = None,              # X軸表示形式 ("0.00", "0.0E+00", "0%" など), Noneで変更なし
                x_log: bool|None = None,      # Trueで対数表示, Noneで変更なし
                y_title: str|bool|None = None,# Y軸タイトル, Noneで変更なし
                y_title_space = +0,           
                y_min: float|str|None = None,  
                y_max: float|str|None = None,  
                y_major: float|None = None,    
                y_minor: float|None = None,    
                y_cross = None,                    # X軸との交差位置, Noneで変更なし        
                y_format = None,              
                y_log: bool|None = None,       
                y2_title: str|bool|None = None,    # Y軸タイトル文字列, Falseで無効化, Noneで変更なし
                y2_min: float|str|None = None,
                y2_max: float|str|None = None,
                y2_major: float|None = None,
                y2_minor: float|None = None,
                y2_format = None,
                y2_log: bool|None = None,       
                frame_color: bool|int|None = None, # グラフ枠色 (False:枠なし, 0:黒枠, Noneでpreset)
                legend: bool|str|None = None,      # 凡例表示(False:非表示, Noneで変更なし, "right"で右側に表示) 
                legend_font_size = None,           # 凡例のフォントサイズ, Noneで変更なし
                legend_width_inc = 0,              # 凡例ボックスの幅増減(pt)
                legend_height_inc = 0,             # 凡例ボックスの高さ増減(pt)
                legend_right_space = 0,            # 凡例の右端 = プロットエリアの右端を基準とした凡例位置制御
                legent_top_space = 0,              # 凡例の上端 = プロットエリアの上端を基準とした凡例位置制御
                transparent_bg: bool|None = None,  # 背景を透明化する場合はTrue
                chart_type = None,                 # "bar"で棒グラフに変更
                x_bold_line: float|None = None,    # x_bold_line=0でx=0が太線
                y_bold_line: float|None = None,    # y_bold_line=0でy=0が太線
                plot_area_space: str = "relative", # プロットエリア調整時の基準位置を"absolute"/"relative"で指定
                width_inc = 0,                     # プロットエリアの幅増減(pt)
                height_inc = 0,                    # プロットエリアの高さ増減(pt)
                ):

    p = PRESET.get(preset, PRESET["std"]) or {}
    
    # 注意) series_listの系列数をデータ範囲以下にしないと例外発生となる
    # list / dict はミュータブルのため、デフォルト引数は None
    if series_list is None:
        series_list = [{"color": "blue"}]
 
    COLOR_NAME_TO_RGB = {
                    "blue": RGB(68,114,196),
                    "orange":RGB(237,125,49),
                    "green":RGB(112,173,71),
                    "yellow":RGB(255,192,0),
                    "purple":RGB(112,48,160),
                    "gray":RGB(165,165,165)
                    }       
    for cfg in series_list:
        c = cfg.get("color")
        if isinstance(c, str):
            cfg["color"] = COLOR_NAME_TO_RGB.get(c.lower(), None)
            
    # グラフ全体のサイズ変更
    if width_cm not in (None, "", 0):
        chart.width = cm_to_pt(width_cm)
    if height_cm not in (None, "", 0):
        chart.height = cm_to_pt(height_cm)  

    # グラフのチャート名 (エクセル画面左上の表示で確認できる)
    if name not in (None, ""):
        chart.name = name

    # Chart本体 → api[1]
    ch = chart.api[1]

    # グラフタイトル
    if title == False:
        ch.HasTitle = True
        ch.ChartTitle.Text = ""
    elif title not in (None, ""):
        ch.HasTitle = True
        ch.ChartTitle.Text = title

    # 横軸のオプション
    x_axis = ch.Axes(AxisType.xlCategory)
    if x_min == "auto":              
        x_axis.MinimumScaleIsAuto = True
    elif x_min not in (None, ""):
        x_axis.MinimumScale = x_min  # 最小値
    if x_max == "auto":               
        x_axis.MaximumScaleIsAuto = True
    elif x_max not in (None, ""):
        x_axis.MaximumScale = x_max  # 最大値
    x_axis.HasMajorGridlines = p.get("x_major_grid", True)
    x_axis.MajorTickMark = p.get("x_major_tickmark", constants.xlTickMarkNone)
    if p.get("x_minor_grid") is not None:
        x_axis.HasMinorGridlines = p.get("x_minor_grid")
    if p.get("x_minor_tickmark") is not None:
        x_axis.MinorTickMark = p.get("x_minor_tickmark") 
    if x_major not in (None, ""):
        x_axis.MajorUnit = x_major   # 目盛間隔          
    if x_minor not in (None, ""):
        x_axis.MinorUnit = x_minor        
    if x_cross not in (None, ""):    # 交差位置(縦軸との交点)
        x_axis.CrossesAt = x_cross
    if x_format not in (None, ""): 
        x_axis.TickLabels.NumberFormatLocal = x_format
    if x_log not in (None, ""):
        x_axis.ScaleType = 1 if x_log == True else 0

    # 横軸のタイトル
    if x_title == False:
        x_axis.HasTitle = False
    elif x_title not in (None, ""):
        x_axis.HasTitle = True
        x_axis.AxisTitle.Text = x_title

    # 縦軸のオプション
    y_axis = ch.Axes(AxisType.xlValue)
    if y_min == "auto":               
        y_axis.MinimumScaleIsAuto = True
    elif y_min not in (None, ""):
        y_axis.MinimumScale = y_min
    if y_max == "auto":               
        y_axis.MaximumScaleIsAuto = True
    elif y_max not in (None, ""):
        y_axis.MaximumScale = y_max
    y_axis.HasMajorGridlines = p.get("y_major_grid", True)
    y_axis.MajorTickMark = p.get("y_major_tickmark", constants.xlTickMarkNone)
    if p.get("y_minor_grid") is not None:
        y_axis.HasMinorGridlines = p.get("y_minor_grid")
    if p.get("y_minor_tickmark") is not None:
        y_axis.MinorTickMark = p.get("y_minor_tickmark")
    if y_major not in (None, ""):
        y_axis.MajorUnit = y_major   # 目盛間隔
    if y_minor not in (None, ""):
        y_axis.MinorUnit = y_minor
    if y_cross not in (None, ""):    # 交差位置(縦軸との交点)
        y_axis.CrossesAt = y_cross
    if y_format not in (None, ""): 
        y_axis.TickLabels.NumberFormatLocal = y_format
    if y_log not in (None, ""):
        y_axis.ScaleType = 1 if y_log == True else 0

    # 縦軸のタイトル
    if y_title == False:
        y_axis.HasTitle = False
    elif y_title not in (None, ""):
        y_axis.HasTitle = True
        y_axis.AxisTitle.Text = y_title   
    if chart_type not in (None,""):
        if chart_type == "bar":
            chart_type = "column_clustered"
        chart.chart_type = chart_type
    
    # 系列の設定 ----------------------------------------------------------------------------- 
    use_secondary = False
    NS = max(len(series_list), NS)
    for i in range(1, NS + 1):
        if i <= len(series_list):
            cfg = series_list[i - 1]
        else:
            cfg = {}
        try:
            series = ch.SeriesCollection(i)
        except Exception as e:
            print(f"系列{i}を取得できません: {e}")
            continue
            
        # 系列名
        if cfg.get("name") not in (None, ""):
            series.Name = cfg["name"]
            
        # XValues / Values
        if cfg.get("XValues") not in (None, ""):
            try:
                wsx = cfg.get("sheet",ws)
                if isinstance(wsx, str):
                    wsx = ws.parent.sheets[wsx]
                series.XValues = wsx.range(cfg["XValues"]).api
            except Exception:
                print(f"系列{i}: wsの設定または範囲指定に問題")
        if cfg.get("Values") not in (None, ""):
            try:
                wsy = cfg.get("sheet",ws)
                if isinstance(wsy, str):
                    wsy = ws.parent.sheets[wsy]
                series.Values  = wsy.range(cfg["Values"]).api
            except Exception:
                print(f"系列{i}: wsの設定または範囲指定に問題")     
        
        try:
            if cfg.get("chart_type") == "bar":
                series.ChartType = constants.xlColumnClustered

            # 色
            color = cfg.get("color")
            if color not in (None, ""):
                try:
                    series.Format.Line.ForeColor.RGB = color    # 線の色
                    series.MarkerForegroundColor = color        # マーカー枠線の色
                    series.MarkerBackgroundColor = color        # マーカー内部の色
                except:
                    # hasattr(series.Format, "Fill")
                    series.Format.Fill.ForeColor.RGB = color

            # スムーズ
            if cfg.get("smooth") not in (None, ""):
                smooth_i = cfg["smooth"]
            elif smooth not in (None, ""):
                smooth_i = smooth
            else:
                smooth_i = p.get("smooth",True)
            if smooth_i not in (None, ""):
                try:
                    series.Smooth = bool(smooth_i)
                except:
                    pass

            # スタイル(線とマーカー)
            if cfg.get("style") not in (None, ""):
                style_i = cfg["style"]
            elif style not in (None, ""):
                style_i = style
            else:
                style_i = p.get("style","line+marker") or ""
                         
            if "marker" in style_i:
                if cfg.get("marker") not in (None, ""):
                    marker_i = cfg["marker"]
                elif marker not in (None, ""):
                    marker_i = marker
                elif p.get("marker") not in (None, ""):
                    marker_i = p["marker"]
                else:
                    marker_i = None
                series.MarkerStyle = marker_map.get(marker_i, constants.xlMarkerStyleCircle)  # マーカー:〇
                series.MarkerSize = cfg.get("size",p.get("marker_size"))                      # マーカーサイズ
            else:
                series.MarkerStyle = constants.xlMarkerStyleNone
            if "line" in style_i:
                series.Format.Line.Visible = True
                series.Format.Line.Weight = cfg.get("weight", p.get("line_weight"))  # 線の太さ(pt)
            elif isinstance(style_i, str) and style_i.startswith("dash"):
                series.Format.Line.Visible = True
                series.Format.Line.DashStyle = 4
            elif isinstance(style_i, str) and style_i.startswith("chain"):
                series.Format.Line.Visible = True
                series.Format.Line.DashStyle = 5
            else:
                series.Format.Line.Visible = False
            
            # 線の透明度の設定(0~1), デフォルト値は0
            if cfg.get("alpha") not in (None, ""):
                alpha_i = cfg["alpha"]
            elif alpha not in (None, ""):
                alpha_i = alpha
            elif p.get("alpha") not in (None, ""):
                alpha_i = p["alpha"]
            else:
                alpha_i = None
            if alpha_i not in (None, ""):
                series.Format.Line.Transparency = float(alpha_i) 
                
            axis_name = cfg.get("axis", "primary")
            if axis_name == "secondary" or axis_name == "y2":
                series.AxisGroup = constants.xlSecondary
                use_secondary = True
            else:
                series.AxisGroup = constants.xlPrimary
            
            tl = cfg.get("trendline")
            if tl not in (None, ""):
                try:
                    series.Trendlines().Delete()
                except Exception:
                    pass
                trend = None
                if tl == 0:
                    pass
                elif isinstance(tl, int):
                    if tl == 1:
                        trend = series.Trendlines().Add(Type=constants.xlLinear)
                    elif 2 <= tl <= 6:
                        trend = series.Trendlines().Add(Type=constants.xlPolynomial)
                        trend.Order = tl
                elif isinstance(tl, str):
                    tl_map = {
                        "exp": constants.xlExponential,   # 指数
                        "log": constants.xlLogarithmic,   # 対数
                        "pow": constants.xlPower,         # 累乗
                        "mov": constants.xlMovingAvg,     # 移動平均
                    }
                    ttype = tl_map.get(tl.lower())
                    if ttype is not None:
                        trend = series.Trendlines().Add(Type=ttype)
                if trend is not None:
                    if cfg.get("trendline_name") is not None:
                        trend.Name = cfg.get("trendline_name", "Regression line")
                    trend.Format.Line.ForeColor.RGB = cfg.get("trendline_color",cfg.get("color",RGB(0,0,0)))
                    trend.Format.Line.Weight = cfg.get("trendline_weight", 1.5)
                    Dashstyle = cfg.get("trendline_style", "solid").lower()
                    if Dashstyle == "dash":
                        trend.Format.Line.DashStyle = 4
                    else:  # solid / default
                        # trend.Format.Line.DashStyle = constants.xlSolid
                        pass
                    t_option = cfg.get("trendline_option", "")
                    if "eq" in t_option: 
                        trend.DisplayEquation = True # 近似式を表示
                        # trend.DataLabel.Left
                        # trend.DataLabel.Top
                    if "r2" in t_option: 
                        trend.DisplayRSquared = True # 決定係数(R²)を表示
            
            if cfg.get("legend") is False:
                series.HasLegendKey = False
                
            if cfg.get("data_label", False):
                series.HasDataLabels = True

        except Exception as e:
            print(f"系列{i}の設定でエラー発生:{e}")
    
    # フォーマットの設定 -------------------------------------------------
    try:
        if ch.HasTitle:
            ch.ChartTitle.Format.TextFrame2.TextRange.Font.Bold = p.get("title_font_bold", False)
            ch.ChartTitle.Format.TextFrame2.TextRange.Font.Name = p.get("title_font_name", "Aptos Narrow 本文")
            if title_font_color not in (None, ""):
                ch.ChartTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = title_font_color
            else:
                ch.ChartTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = p.get("title_font_color", RGB(89,89,89))
            if title_font_size not in (None, ""):
                ch.ChartTitle.Format.TextFrame2.TextRange.Font.Size = title_font_size
            else:
                ch.ChartTitle.Format.TextFrame2.TextRange.Font.Size = p.get("title_font_size", 14)
    except Exception as e:
        print("タイトル設定でエラー:",e) 
        
    try:   
        # グリッド線の設定（デフォルト: 薄いグレー)
        axes = [x_axis, y_axis]
        for ax in axes:
            if ax.HasMajorGridlines:
                ax.MajorGridlines.Format.Line.ForeColor.RGB = p.get("major_grid_color", RGB(217, 217, 217))
                ax.MajorGridlines.Format.Line.Weight = p.get("major_grid_weight", 0.75)
            if p.get("axis_line") is None:
                pass
            elif p.get("axis_line") is True:
                ax.Format.Line.Visible = True
                ax.Format.Line.ForeColor.RGB = p.get("axis_line_color", RGB(191, 191, 191))
                ax.Format.Line.Weight = p.get("axis_line_weight", 0.75)
            elif p.get("axis_line") is False:
                ax.Format.Line.Visible = False
    except Exception as e:
        print("軸の設定でエラー:",e)
        
    try:
        # 副軸の設定
        if use_secondary:
            y2 = ch.Axes(AxisType.xlValue, constants.xlSecondary)
            axes.append(y2)
            if y2_min == "auto":                
                y2.MinimumScaleIsAuto = True
            elif y2_min not in (None, ""):
                y2.MinimumScale = y2_min
            if y2_max == "auto":                
                y2.MaximumScaleIsAuto = True
            elif y2_max not in (None, ""):
                y2.MaximumScale = y2_max
            y2.HasMajorGridlines = p.get("y2_major_grid",False)
            y2.MajorTickMark = p.get("y2_major_tickmark", constants.xlTickMarkNone)
            if p.get("y2_minor_grid") is not None:
                y2.HasMinorGridlines = p.get("y2_minor_grid")
            if p.get("y2_minor_tickmark") is not None:
                y2.MinorTickMark = p.get("y2_minor_tickmark")
            if y2_major not in (None, ""):    
                y2.MajorUnit = y2_major
            if y2_minor not in (None, ""):
                y2.MinorUnit = y2_minor
            if y2_format not in (None, ""):
                y2.TickLabels.NumberFormatLocal = y2_format
            if y2_log not in (None, ""):
                y2.ScaleType = 1 if y2_log == True else 0
            if y2_title == False:
                y2.HasTitle = False
            elif y2_title not in (None, ""):
                y2.HasTitle = True
                y2.AxisTitle.Text = y2_title
    except Exception as e:
        print("副軸の設定でエラー:",e)
        
    try:    
        for ax in axes: 
            # 軸の設定
            tl_font = ax.TickLabels.Font
            tl_font.Color = p.get("axis_tick_font_color", RGB(89,89,89))
            tl_font.Size = p.get("axis_tick_font_size", 9)
            tl_font.Name = p.get("axis_tick_font_name","Aptos Narrow 本文")
            # 軸タイトルがあるとき、軸タイトルを設定する
            if ax.HasTitle:
                ax_font = ax.AxisTitle.Format.TextFrame2.TextRange.Font
                ax_font.Fill.ForeColor.RGB = p.get("axis_title_font_color", RGB(89,89,89))
                ax_font.Bold = p.get("axis_title_font_bold", False)
                ax_font.Size = p.get("axis_title_font_size", 10)
                ax_font.Name = p.get("axis_title_font_name", "Aptos Narrow 本文")
    except Exception as e:
        print("軸のフォント設定でエラー:",e)
        
    try:
        # chart_area/plot_area のオブジェクト取得
        chart_area = ch.ChartArea
        plot_area = ch.PlotArea
        
        # 外枠の設定
        if frame_color is False:  # False:枠なし、0:黒枠
            chart_area.Border.LineStyle = 0    # 枠なし
        elif frame_color not in (None, ""):
            chart_area.Format.Line.ForeColor.RGB = frame_color
            chart_area.Format.Line.Weight = p.get("frame_weight",0.75)
        else:
            chart_area.Format.Line.ForeColor.RGB = p.get("frame_color",RGB(217,217,217))  # 薄いグレー
            chart_area.Format.Line.Weight = p.get("frame_weight",0.75)                    # 枠線の太さ(pt)
        
        # プロットエリアの枠設定
        if p.get("plot_area_frame") is not None:
            plot_area.Format.Line.Visible = p.get("plot_area_frame", False)
            if p.get("plot_area_frame") is True:
                plot_area.Format.Line.ForeColor.RGB = p.get("plot_area_frame_color", 0)
                plot_area.Format.Line.Weight = p.get("plot_area_frame_weight", 1.0)

        # 背景の透明化設定
        if transparent_bg is True:
            chart_area.Format.Fill.Visible = False
            plot_area.Format.Fill.Visible = False
        elif transparent_bg is False:
            chart_area.Format.Fill.Visible = True
            plot_area.Format.Fill.Visible = True
            chart_area.Format.Fill.ForeColor.RGB = RGB(255,255,255)
            plot_area.Format.Fill.ForeColor.RGB = RGB(255,255,255)
            
        # ラインの強調
        if x_bold_line is not None:
            emphasize_line(x_axis,x_bold_line)
        if y_bold_line is not None:
            emphasize_line(y_axis,y_bold_line)
    except Exception as e:
        print("チャートエリア、プロットエリアの設定でエラー:",e)

    try: 
        # 凡例を一度無効にする(例外あり)
        if legend == "right":
            ch.HasLegend = True
        elif legend not in (None, ""):
            ch.HasLegend = False
        
        # プロットエリアの調整
        if plot_area_space in ("absolute","abs"):
            plot_area.InsideLeft   = chart_area.Left + y_title_space
            plot_area.InsideTop    = chart_area.Top + title_space
            plot_area.InsideWidth  = chart_area.Width - y_title_space + width_inc
            plot_area.InsideHeight = chart_area.Height - title_space - x_title_space + height_inc
        else:  # relative / dafault
            plot_area.InsideLeft   = plot_area.InsideLeft + y_title_space
            plot_area.InsideTop    = plot_area.InsideTop + title_space
            plot_area.InsideWidth  = plot_area.InsideWidth - y_title_space + width_inc
            plot_area.InsideHeight = plot_area.InsideHeight - title_space - x_title_space + height_inc
        
        # 凡例設定
        if  legend in (None, ""):
            pass
        elif legend == False:
            ch.HasLegend = False
        else:
            ch.HasLegend = True
            if legend == "auto":
                pass
            else:
                if legend_font_size not in (None, ""):
                    ch.Legend.Format.TextFrame2.TextRange.Font.Size = legend_font_size
                if legend_width_inc!=0:
                    ch.Legend.Width  += legend_width_inc
                if legend_height_inc!=0:
                    ch.Legend.Height += legend_height_inc
                if "T" in legend:   # 大文字のTを含む場合
                    ch.Legend.Top = plot_area.InsideTop + legent_top_space
                elif "B" in legend: # 大文字のBを含む場合
                    ch.Legend.Top = plot_area.InsideTop + plot_area.InsideHeight - ch.Legend.Height - legent_top_space
                if "R" in legend:   # 大文字のRを含む場合
                    ch.Legend.Left = plot_area.InsideLeft + plot_area.InsideWidth - ch.Legend.Width - legend_right_space
                elif "L" in legend: # 大文字のLを含む場合
                    ch.Legend.Left = plot_area.InsideLeft + legend_right_space
                elif "C" in legend: # 大文字のCを含む場合
                    ch.Legend.Left = plot_area.InsideLeft + (plot_area.InsideWidth - ch.Legend.Width)/2
                if "bw" in legend: # background white
                    ch.Legend.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
                    ch.Legend.Format.Fill.Visible = True
                if "fb" in legend: # frame black
                    ch.Legend.Format.Line.ForeColor.RGB = RGB(0, 0, 0)
                    ch.Legend.Format.Line.Weight = 0.75
                    ch.Legend.Format.Line.Visible = True
                if "tb" in legend: # text black
                    ch.Legend.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    except Exception as e:
        print("凡例 もしくは プロットエリアの調整でエラー:",e)

    return chart

# プロットエリア内に帯や注釈用のボックスを配置する
def add_shape(chart, x_start=None, x_end=None, y_start=None, y_end=None, 
              left=None,  width=None, top=None, height=None, delta_width=0, delta_height=0, 
              color=RGB(0,0,255), alpha=0.8, frame_color=None, white_box=None,
              text=None, font_color=RGB(0,0,0), font_name="Calibri", font_size=10,
              font_bold=True, alignment=None, auto_size=None,
              right=None, bottom=None):
    """
    優先順位  帯の指定(x_start, x_end) > 位置の指定(left, right) の両者指定 > left or right と width 
                     (y_start, y_end) > 位置の指定(top, bottom) の両者指定 > top or bottom と height
    1) start, x_end, y_start, y_endをすべてNoneとすることで位置 (数値pt または "数値%") で指定可能です。
    2) left, right, top, bottomには 負の数値、"負の数値%" を指定できます。width,heightは0より大きい値を指定してください。
    3) left + right > width もしくは top + bottom > height では負のサイズにより、エラー発生します。
    4) delta_width, delta_heightで Excel描画都合の微調整ができます。
    5) white_box = True では、不透明の白背景、黒枠のテキストボックスとなります。それ以外の条件はNoneにしてください。
    6) white_boxのNone/True に関わらず、text = "1行目\n2行目" (改行\n) でテキスト文字を指定できます。
    7) テキストありで、auto_size=True の場合、left\topを基準に右下方向へ自動拡張されます。
    8) alignment ではテキスト配置を 左揃え:"left"、中央揃え:"center"、右揃え:"right" 
                                   と 上揃え:"top", 上下中央揃え:"middle", 下揃え:"bottom を組み合わせて指定できます。
    """
    ch = chart.api[1]
    plot_area = ch.PlotArea
    x_axis = ch.Axes(constants.xlCategory)
    x_min, x_max = x_axis.MinimumScale, x_axis.MaximumScale
    y_axis = ch.Axes(constants.xlValue)
    y_min, y_max = y_axis.MinimumScale, y_axis.MaximumScale
      
    # (x_start=None, x_end=None) or (y_start=None, y_end=None) で使用する関数
    def parse_size(value, total_size):
        """
        value:
            - 数値 → px
            - "xx%" → total_size * xx/100
            - None → 呼び出し側で解釈
        """
        if isinstance(value, str) and value.endswith("%"):
            try:
                pct = float(value.strip("%")) / 100
                return total_size * pct
            except ValueError:
                raise ValueError(f"Invalid percentage value: {value}")
        else:
            return value
        
    def resolve_1d(start, end, axis_min, axis_max, left, right, width, origin, total):
        if start is not None and end is None: end   = axis_max
        if start is None and end is not None: start = axis_min
        if start is not None and end is not None:
            left = origin + (start - axis_min) / (axis_max - axis_min) * total
            width = (end - start) / (axis_max - axis_min) * total
        elif left is None and right is None:
            left = origin
            width = parse_size(width, total)
            if width is None:
                width = total
        elif left is not None and right is not None:
            left_val  = parse_size(left,  total)
            right_val = parse_size(right, total)
            left  = origin + left_val
            width = total - left_val - right_val
        elif left is not None and right is None:
            left_val = parse_size(left, total)
            left = origin + left_val
            width = parse_size(width, total)
            if width is None:
                width = total - left_val
        elif left is None and right is not None:
            right_val = parse_size(right, total)
            width = parse_size(width, total)
            if width is None:
                width = total - right_val
            left = origin + total - width - right_val
        return left, width
    
    # X軸座標変換
    origin = plot_area.InsideLeft
    total  = plot_area.InsideWidth  
    left, width = resolve_1d(x_start, x_end, x_min, x_max, left, right, width, origin, total)
    # Excel描画の都合の微調整
    width += delta_width
    # Y軸座標変換
    origin = plot_area.InsideTop
    total  = plot_area.InsideHeight
    top, height = resolve_1d(y_start, y_end, y_min, y_max, top, bottom, height, origin, total)
    # Excel描画の都合の微調整
    height += delta_height
    
    if width <= 0 or height <= 0:
        raise ValueError(f"Invalid shape size: width={width}, height={height}")
    
    if white_box:
        shape = ch.Shapes.AddTextbox(1, left, top, width, height)
        color = RGB(255,255,255)
        alpha = 0
        frame_color = RGB(0,0,0)
    else:
        # 1 = msoShapeRectangle
        shape = ch.Shapes.AddShape(1, left, top, width, height)
             
    shape.Fill.ForeColor.RGB = color
    shape.Fill.Transparency = alpha
    if frame_color is None:
        shape.Line.Visible = False
    else:
        shape.Line.Visible = True
        shape.Line.ForeColor.RGB = frame_color
        shape.Line.Weight = 0.75
    if text is not None:
        shape.TextFrame.Characters().Text = text
        if auto_size:
            shape.TextFrame.AutoSize = True 
        shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = font_color
        shape.TextFrame2.TextRange.Font.Name = font_name
        shape.TextFrame2.TextRange.Font.Size = font_size
        shape.TextFrame2.TextRange.Font.Bold = font_bold
        if alignment is not None:
            alignment = alignment.strip().lower()
            if "left" in alignment:
                shape.TextFrame.HorizontalAlignment = constants.xlHAlignLeft  
            elif "center" in alignment:
                shape.TextFrame.HorizontalAlignment = constants.xlHAlignCenter
            elif "right" in alignment:
                shape.TextFrame.HorizontalAlignment = constants.xlHAlignRight
            if "top" in alignment:
                shape.TextFrame.VerticalAlignment = constants.xlVAlignTop
            elif "middle" in alignment:
                shape.TextFrame.VerticalAlignment   = constants.xlVAlignCenter
            elif "bottom" in alignment:
                shape.TextFrame.VerticalAlignment = constants.xlVAlignBottom
    return shape
                
def add_line(chart, x=None, y=None, color=RGB(0, 0, 0), weight=1.5, dash=True):
    """
    プロットエリア内に基準線を描画する

    x : 縦線（x軸の値）
    y : 横線（y軸の値）
    """

    if (x is None) == (y is None):
        raise ValueError("x または y のどちらか一方だけ指定してください")

    ch = chart.api[1]
    plot = ch.PlotArea

    x_axis = ch.Axes(constants.xlCategory)
    y_axis = ch.Axes(constants.xlValue)

    # --- 縦線 ---
    if x is not None:
        x_pos = plot.InsideLeft + (x - x_axis.MinimumScale)/(x_axis.MaximumScale - x_axis.MinimumScale)*plot.InsideWidth
        x1 = x2 = x_pos
        y1 = plot.InsideTop
        y2 = plot.InsideTop + plot.InsideHeight

    # --- 横線 ---
    else:
        y_pos = plot.InsideTop + (1 - (y - y_axis.MinimumScale)/(y_axis.MaximumScale - y_axis.MinimumScale))*plot.InsideHeight
        y1 = y2 = y_pos
        x1 = plot.InsideLeft
        x2 = plot.InsideLeft + plot.InsideWidth

    line = ch.Shapes.AddLine(x1, y1, x2, y2)
    line.Line.ForeColor.RGB = color
    line.Line.Weight = weight
    if dash:
        line.Line.DashStyle = 4       
    return line