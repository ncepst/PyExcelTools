# ScatterChart.py
import xlwings as xw
from xlwings.constants import AxisType
from win32com.client import constants

# 新規グラフを作成します
# from ScatterChart import ScatterChart, RGB
# ScatterChart() で呼び出し

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
        "title_font_size":14,
        "title_font_color":RGB(89,89,89),
        "title_font_bold":False,
        "title_font_name":"Aptos Narrow 本文",
        "axis_title_font_size":10,
        "axis_title_font_name":"Aptos Narrow 本文",
        "axis_title_font_color":RGB(89,89,89),
        "axis_title_font_bold":False,
        "axis_tick_font_color": RGB(89, 89, 89),
        "axis_tick_font_size": 9,
        "axis_tick_font_name":"Aptos Narrow 本文", 
        "axis_line_color":RGB(191, 191, 191),        
        "x_major_grid": True,
        "y_major_grid": True,
        "x_minor_grid": False,
        "y_minor_grid": False,
        "major_grid_color": RGB(217, 217, 217),
        "major_grid_weight":0.75,
        # TickMark: None, Inside, Outside, Cross
        "major_tickmark":constants.xlTickMarkNone,  # 目盛の内向き/外向きなし
        "frame_color":RGB(217,217,217),             # False:枠なし
        "frame_weight":0.75,
        "style":"line+marker",
        "smooth":True,
        "alpha":None,
        "line_weight":1.5,
        "marker": "C",
        "marker_size":5,
        "y2_major_grid":False,  # 副軸グリッド表示なし
        "y2_minor_grid":False,
        "y2_major_tickmark":constants.xlTickMarkNone,
    },
    "std": {
        "axis_title_font_color":RGB(0,0,0), 
        "axis_tick_font_color": RGB(0,0,0),
    },
}
# std を excel2021 ベースで上書き
PRESET["std"] = {**PRESET["excel2021"], **PRESET["std"]}

marker_map = {
        "C":constants.xlMarkerStyleCircle,
        "S":constants.xlMarkerStyleSquare,
        "D":constants.xlMarkerStyleDiamond,
        "T":constants.xlMarkerStyleTriangle,
        "N":constants.xlMarkerStyleNone
        }

# titleやframe_colorには、False(無効化)もしくは文字・値を入力します
# 優先順位 series cfg > 引数 > preset
def ScatterChart(ws,
                 start_range = "A1",
                 row = None,
                 col = None,
                 paste_range = "A1",
                 width_cm = 12.54,
                 height_cm = 7.54,
                 preset = "std",
                 name = None,
                 title = False,       # 無効化:False
                 title_font_color = None,
                 title_font_size = None,
                 title_space = +0,
                 series_list = None,
                 style= None,
                 smooth = None,
                 marker = None,
                 alpha = None,
                 x_title = False,     # 無効化:False
                 x_title_space = -15, # 下に広げる場合はマイナス
                 x_min = None,
                 x_max = None,
                 x_major = None,
                 x_minor = None,
                 x_cross = None,
                 x_format = None,
                 x_log = None,
                 y_title = False,     # 無効化:False
                 y_title_space = +0,
                 y_min = None,
                 y_max = None,
                 y_major = None,
                 y_minor = None,
                 y_cross = None,
                 y_format = None,
                 y_log = None,
                 y2_title = False,
                 y2_min = None,
                 y2_max = None,
                 y2_major = None,
                 y2_minor = None,
                 y2_format = None,
                 y2_log = None,
                 frame_color = None,  # 枠なし:False, 黒枠:0
                 width_inc = 0,
                 height_inc = 0,
                 legend = False,      # 無効化:False
                 legend_font_size = None,
                 legend_width_inc = 0,
                 legend_height_inc = 0,
                 legend_right_space = 0,
                 transparent_bg = None,
                 chart_type = None, 
                 x_bold_line = None,
                 y_bold_line = None,
                ):
    
    p = PRESET.get(preset, PRESET["std"]) or {}
    
    # list / dict はミュータブルのため、デフォルト引数を None にしている
    if series_list is None:
        series_list = [{"color": RGB(68,114,196)}]
    """
    注意) series_listの系列数をデータ範囲以下にしないと例外発生となる
    series_list = [{"name":"系列1", "color": RGB(68,114,196)},  # 青
                   {"name":"系列2", "color": RGB(237,125,49)},  # オレンジ
                   {"name":"系列3", "color": RGB(112,173,71)},  # 緑
                   {"name":"系列4", "color": RGB(165,165,165)}, # グレー
                  ],
    """                
        
    # ----------------------------------------------------------
    # 散布図のエクセルグラフを作成する
    # ----------------------------------------------------------
    # (セル範囲入力) --------------------------------------------
    start_range = start_range
    start = ws[start_range]
    # row, col が None の場合は自動で連続データ範囲を検出
    if row is None:
        row = start.end('down').row - start.row + 1
    if col is None:
        col = start.end('right').column - start.column + 1
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
    if name not in (None, ""):
        chart.name = name

    # Chart本体 → api[1]
    ch = chart.api[1]
    
    # グラフタイトル
    if title in (False, ""):
        ch.HasTitle = True
        ch.ChartTitle.Text = ""
    else:
        ch.HasTitle = True
        if title in (True, None):
            ch.ChartTitle.Text = "グラフタイトル"
        else:
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
    if x_major not in (None, ""):     
        x_axis.MajorUnit = x_major   # 目盛間隔
    x_axis.HasMinorGridlines = p.get("x_minor_grid", False)           
    if x_minor not in (None, ""):
        x_axis.MinorUnit = x_minor         
    if x_cross not in (None, ""):     # 交差位置(縦軸との交点)
        x_axis.CrossesAt = x_cross
    if x_format not in (None, ""): 
        x_axis.TickLabels.NumberFormatLocal = x_format
    if x_log not in (None, ""):
        x_axis.ScaleType = 1 if x_log == True else 0

    # 横軸のタイトル
    if x_title in (False, ""):
        x_axis.HasTitle = False
    else:
        x_axis.HasTitle = True
        if x_title in (True, None):
            x_axis.AxisTitle.Text = "x軸タイトル"
        else:
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
    if y_major not in (None, ""):
        y_axis.MajorUnit = y_major   # 目盛間隔
    y_axis.HasMinorGridlines = p.get("y_minor_grid", False)
    if y_minor not in (None, ""):
        y_axis.MinorUnit = y_minor 
    if y_cross not in (None, ""):     # 交差位置(縦軸との交点)
        y_axis.CrossesAt = y_cross
    if y_format not in (None, ""): 
        y_axis.TickLabels.NumberFormatLocal = y_format
    if y_log not in (None, ""):
        y_axis.ScaleType = 1 if y_log == True else 0

    # 縦軸のタイトル
    if y_title in (False, ""):
        y_axis.HasTitle = False
    else:
        y_axis.HasTitle = True
        if y_title in (True, None):
            y_axis.AxisTitle.Text = "y軸タイトル"
        else:
            y_axis.AxisTitle.Text = y_title
        
    if chart_type not in (None,""):
        if chart_type == "bar":
            chart_type = "column_clustered"
        chart.chart_type = chart_type
    
    # 系列の設定 -----------------------------------------------------------------------------
    use_secondary = False
    NS = min(row, col)-1
    NS = max(len(series_list), NS)
    # print("系列数:", NS)
    for i in range(1, NS + 1):
        if i <= len(series_list):
            cfg = series_list[i - 1]
        else:
            cfg = {}
        try:
            series = ch.SeriesCollection(i)
            
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
                    
            # 色
            color = cfg.get("color")
            if color not in (None, ""):
                series.Format.Line.ForeColor.RGB = color    # 線の色
                series.MarkerForegroundColor = color        # マーカー枠線の色
                series.MarkerBackgroundColor = color        # マーカー内部の色
            
            # スムーズ
            if cfg.get("smooth") not in (None, ""):
                smooth_i = cfg["smooth"]
            elif smooth not in (None, ""):
                smooth_i = smooth
            else:
                smooth_i = p.get("smooth",True)
            if smooth_i not in (None, ""):
                series.Smooth = bool(smooth_i)

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
            if cfg.get("chart_type") == "bar":
                series.ChartType = constants.xlColumnClustered
            
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
            print(f"系列{i}で例外発生:{e}")
    
    # デフォルトはExcel 2021 以降の標準スタイル -------------------------------------------------
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
        
        # グリッド線の設定（デフォルト: 薄いグレー)
        if x_axis.HasMajorGridlines:
            x_axis.MajorGridlines.Format.Line.ForeColor.RGB = p.get("major_grid_color", RGB(217, 217, 217))
            x_axis.MajorGridlines.Format.Line.Weight = p.get("major_grid_weight", 0.75)
        x_axis.Format.Line.ForeColor.RGB = p.get("axis_line_color", RGB(191, 191, 191))
        x_axis.MajorTickMark = p.get("major_tickmark", constants.xlTickMarkNone)  # 目盛の内向き/外向きなし
        if y_axis.HasMajorGridlines:
            y_axis.MajorGridlines.Format.Line.ForeColor.RGB = p.get("major_grid_color", RGB(217, 217, 217))
            y_axis.MajorGridlines.Format.Line.Weight = p.get("major_grid_weight", 0.75)
        y_axis.Format.Line.ForeColor.RGB = p.get("axis_line_color", RGB(191, 191, 191))
        y_axis.MajorTickMark = p.get("major_tickmark", constants.xlTickMarkNone)  # 目盛の内向き/外向きなし
        axes = [x_axis, y_axis]
                
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
            if y2_major not in (None, ""):    
                y2.MajorUnit = y2_major
            y2.HasMinorGridlines = p.get("y2_minor_grid",False) 
            if y2_minor not in (None, ""):
                y2.MinorUnit = y2_minor  
            if y2_format not in (None, ""):
                y2.TickLabels.NumberFormatLocal = y2_format
            if y2_log not in (None, ""):
                y2.ScaleType = 1 if y2_log == True else 0
            if y2_title in (False, ""):
                y2.HasTitle = False
            else:
                y2.HasTitle = True
                if y2_title in (True, None):
                    y2.AxisTitle.Text = "副軸タイトル"
                else:
                    y2.AxisTitle.Text = y2_title
                    
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
                
        # 外枠の設定
        if frame_color is False:  # False:枠なし、0:黒枠
            ch.ChartArea.Border.LineStyle = 0 # 枠なし
        elif frame_color not in (None, ""):
            ch.ChartArea.Format.Line.ForeColor.RGB = frame_color
            ch.ChartArea.Format.Line.Weight = p.get("frame_weight",0.75)
        else:
            ch.ChartArea.Format.Line.ForeColor.RGB = p.get("frame_color",RGB(217,217,217))  # 薄いグレー
            ch.ChartArea.Format.Line.Weight = p.get("frame_weight",0.75)                    # 枠線の太さ(pt)                  
        
        # 背景の透明化設定
        if transparent_bg is True:
            ch.ChartArea.Format.Fill.Visible = False
            ch.PlotArea.Format.Fill.Visible = False
        elif transparent_bg is False:
            ch.ChartArea.Format.Fill.Visible = True
            ch.PlotArea.Format.Fill.Visible = True
            ch.ChartArea.Format.Fill.ForeColor.RGB = RGB(255,255,255)
            ch.PlotArea.Format.Fill.ForeColor.RGB = RGB(255,255,255)
            
        # ラインの強調
        if x_bold_line is not None:
            emphasize_line(x_axis,x_bold_line)
        if y_bold_line is not None:
            emphasize_line(y_axis,y_bold_line)
        # ----------------------------------------------------------------------
    except Exception as e:
        print("フォーマットの設定でエラー:",e)
        
    try: 
        # 凡例を一度無効にする(例外あり)
        if legend == "right":
            ch.HasLegend = True
        else:
            ch.HasLegend = False
        
        # プロットエリアの調整
        pa = ch.PlotArea
        pa.InsideLeft   = pa.InsideLeft + y_title_space
        pa.InsideTop    = pa.InsideTop + title_space
        pa.InsideWidth  = pa.InsideWidth - y_title_space + width_inc
        pa.InsideHeight = pa.InsideHeight - title_space - x_title_space + height_inc
        
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
                if "U" in legend: # 大文字のUを含む場合
                    ch.Legend.Top = ch.PlotArea.InsideTop
                if "R" in legend: # 大文字のRを含む場合
                    ch.Legend.Left = ch.PlotArea.InsideLeft + ch.PlotArea.InsideWidth - ch.Legend.Width - legend_right_space
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