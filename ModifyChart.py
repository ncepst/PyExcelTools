#ModifyChart.py
import xlwings as xw
from xlwings.constants import AxisType
from win32com.client import constants

# 既存グラフの変更
# from ModifyChart import ModifyChart, RGB

def RGB(r, g, b):
    return r + g*256 + b*65536

PRESET = {
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
        "axis_tick_font_name":"Aptos Narrow 本文", # "Aptos Narrow 本文"は Excel 2021以降のみ
        "axis_line_color":RGB(191, 191, 191),        
        "major_grid": True,
        "major_grid_color": RGB(217, 217, 217),
        "major_grid_weight":0.75,
        "major_tickmark":constants.xlTickMarkNone,  # 目盛の内向き/外向きなし
        "frame_color":RGB(217,217,217),             # False:枠なし
        "frame_weight":0.75,
        "style":"line+marker",
        "smooth":True,
        "line_weight":1.5,
        "marker": "C",
        "marker_size":5,
    },
    "std": {
        "axis_title_font_color":RGB(0,0,0), # 黒に変更
        "axis_tick_font_color": RGB(0,0,0), # 黒に変更
    },
}
# std を excel2021 ベースで展開
PRESET["std"] = {**PRESET["excel2021"], **PRESET["std"]}

# False:無効化, None: 変更なし もしくは デフォルト値
def ModifyChart(chart,
                ws = None,
                preset = "std",
                width_cm = None,
                height_cm = None,
                name = None,
                title = None,     # 無効化:False
                title_font_color = None,
                title_font_size = None,
                title_space = +0,
                NS = 1,
                series_list = None,
                style= None,
                smooth = None,
                marker = None,
                alpha = None,
                x_title = None,   # 無効化:False
                x_min = None,
                x_max = None,
                x_major = None,
                x_cross = None,
                x_format = None,
                y_title = None,   # 無効化:False
                x_title_space = +0,
                y_min = None,
                y_max = None,
                y_major = None,
                y_cross = None,
                y_format = None,
                y2_title = None,
                y2_min = None,
                y2_max = None,
                y2_major = None,
                y2_format = None,
                y2_grid = False,   # 副軸はグリッド無し
                frame_color = None,
                y_title_space = +0,
                width_inc = 0,
                height_inc = 0,
                legend = None,     # 無効化:False
                legend_font_size = None,
                legend_width_inc = 0,
                legend_height_inc = 0,
                legend_right_space = 0,
                chart_type = None,
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
    # cm → pt 換算関数の定義 (1 point = 1/72 inch, 1 inch = 2.54 cm)
    def cm_to_pt(cm):
        return cm * 72 / 2.54
    
    # グラフ全体のサイズ変更
    if width_cm not in (None, "", 0):
        chart.width = cm_to_pt(width_cm)
    if height_cm not in (None, "", 0):
        chart.height = cm_to_pt(height_cm)  

    # グラフのチャート名 (エクセル画面左上の表示で確認できる)
    if name not in (None, ""):
        chart.name = name

    # ChartObject(枠) → api[0]、Chart本体 → api[1]
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
    if x_min == "auto":               #最小値
        x_axis.MinimumScaleIsAuto = True
    elif x_min not in (None, ""):
        x_axis.MinimumScale = x_min
    if x_max == "auto":               #最大値
        x_axis.MaximumScaleIsAuto = True
    elif x_max not in (None, ""):
        x_axis.MaximumScale = x_max
    if x_major not in (None, ""):     # 目盛間隔
        x_axis.MajorUnit = x_major    
    if x_cross not in (None, ""):     # 交差位置(縦軸との交点)
        x_axis.CrossesAt = x_cross
    if x_format not in (None, ""): 
        x_axis.TickLabels.NumberFormatLocal = x_format

    # 横軸のタイトル
    if x_title == False:
        x_axis.HasTitle = False
    elif x_title not in (None, ""):
        x_axis.HasTitle = True
        x_axis.AxisTitle.Text = x_title

    # 縦軸のオプション
    y_axis = ch.Axes(AxisType.xlValue)
    if y_min == "auto":               #最小値
        y_axis.MinimumScaleIsAuto = True
    elif y_min not in (None, ""):
        y_axis.MinimumScale = y_min
    if y_max == "auto":               #最大値
        y_axis.MaximumScaleIsAuto = True
    elif y_max not in (None, ""):
        y_axis.MaximumScale = y_max
    if y_major not in (None, ""):     # 目盛間隔
        y_axis.MajorUnit = y_major    
    if y_cross not in (None, ""):     # 交差位置(縦軸との交点)
        y_axis.CrossesAt = y_cross
    if y_format not in (None, ""): 
        y_axis.TickLabels.NumberFormatLocal = y_format

    # 縦軸のタイトル
    if y_title == False:
        y_axis.HasTitle = False
    elif y_title not in (None, ""):
        y_axis.HasTitle = True
        y_axis.AxisTitle.Text = y_title
        
    if chart_type not in (None,""):
        chart.chart_type = chart_type
    
    # 系列の設定 -----------------------------------------------------------------------------
    marker_map = {
            "C":constants.xlMarkerStyleCircle,
            "S":constants.xlMarkerStyleSquare,
            "D":constants.xlMarkerStyleDiamond,
            "T":constants.xlMarkerStyleTriangle
            }
    
    style_all  = style
    smooth_all = smooth
    marker_all = marker
    alpha_all  = alpha
    use_secondary = False
    
    NS = max(len(series_list), NS)
    for i in range(1, NS + 1):
        if i <= len(series_list):
            cfg = series_list[i - 1]
        else:
            cfg = {}
        try:
            series = ch.SeriesCollection(i)
            if cfg.get("name") not in (None, ""):
                series.Name = cfg["name"]
            if cfg.get("XValues"):
                try:
                    series.XValues = ws.range(cfg["XValues"]).api
                except:
                    print("wsの定義が必要です")
            if cfg.get("Values"):
                try:
                    series.Values  = ws.range(cfg["Values"]).api
                except:
                    print("wsの定義が必要です")
            color = cfg.get("color")
            if color not in (None, ""):
                series.Format.Line.ForeColor.RGB = color    # 線の色
                series.MarkerForegroundColor = color        # マーカー枠線の色
                series.MarkerBackgroundColor = color        # マーカー内部の色
            
            if smooth_all in (None, ""):
                smooth = cfg.get("smooth", p.get("smooth"))
            series.Smooth = bool(smooth)
    
            if style_all in (None, ""):
                style = cfg.get("style") or p.get("style") or "line+marker"           
            if "marker" in style:
                if marker_all in (None, ""):
                    marker = cfg.get("marker") or p.get("marker")
                series.MarkerStyle = marker_map.get(marker, constants.xlMarkerStyleCircle)    # マーカー:〇
                series.MarkerSize = cfg.get("size",p.get("marker_size"))                      # マーカーサイズ
            else:
                series.MarkerStyle = constants.xlMarkerStyleNone
            if "line" in style:
                series.Format.Line.Visible = True
                series.Format.Line.Weight = cfg.get("weight", p.get("line_weight"))  # 線の太さ(pt)
            elif isinstance(style, str) and style.startswith("dash"):
                series.Format.Line.Visible = True
                series.Format.Line.DashStyle = 4
            elif isinstance(style, str) and style.startswith("chain"):
                series.Format.Line.Visible = True
                series.Format.Line.DashStyle = 5
            else:
                series.Format.Line.Visible = False
            if alpha_all in (None, ""):   
                alpha = cfg.get("alpha",0) # 透明度は0~1
            if alpha not in (None, ""):
                series.Format.Line.Transparency = float(alpha)
                
            axis_name = cfg.get("axis", "primary")
            if axis_name == "secondary" or axis_name == "y2":
                series.AxisGroup = constants.xlSecondary
                use_secondary = True
            else:
                series.AxisGroup = constants.xlPrimary
            if cfg.get("chart_type") == "bar":
                series.ChartType = constants.xlColumnClustered
                
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
        x_axis.HasMajorGridlines = p.get("major_grid", True)
        if x_axis.HasMajorGridlines:
            x_axis.MajorGridlines.Format.Line.ForeColor.RGB = p.get("major_grid_color", RGB(217, 217, 217))
            x_axis.MajorGridlines.Format.Line.Weight = p.get("major_grid_weight", 0.75)
        x_axis.Format.Line.ForeColor.RGB = p.get("axis_line_color", RGB(191, 191, 191))
        x_axis.MajorTickMark = p.get("major_tickmark", constants.xlTickMarkNone)  # 目盛の内向き/外向きなし
        y_axis.HasMajorGridlines = p.get("major_grid", True)
        if y_axis.HasMajorGridlines:
            y_axis.MajorGridlines.Format.Line.ForeColor.RGB = p.get("major_grid_color", RGB(217, 217, 217))
            y_axis.MajorGridlines.Format.Line.Weight = p.get("major_grid_weight", 0.75)
        y_axis.Format.Line.ForeColor.RGB = p.get("axis_line_color", RGB(191, 191, 191))
        y_axis.MajorTickMark = p.get("major_tickmark", constants.xlTickMarkNone)  # 目盛の内向き/外向きなし

        for ax in ch.Axes():    
            # 軸の設定
            tl_font = ax.TickLabels.Font
            tl_font.Color = p.get("axis_tick_font_color", RGB(89,89,89))
            tl_font.Size = p.get("axis_tick_font_size", 9)
            tl_font.Name = p.get("axis_tick_font_name","Aptos Narrow 本文")
            # 軸タイトルがあるとき、軸タイトルを設定する
            if ax.HasTitle:
                ax_font = ax.AxisTitle.Format.TextFrame2.TextRange.Font
                ax_font.Fill.ForeColor.RGB = p.get("axis_title_font_color", RGB(89,89,89))
                ax_font.TextFrame2.TextRange.Font.Bold = p.get("axis_title_font_bold", False)
                ax_font.Size = p.get("axis_title_font_size", 10)
                ax_font.Name = p.get("axis_title_font_name", "Aptos Narrow 本文")
                
        # 副軸の設定
        if use_secondary:
            y2 = ch.Axes(AxisType.xlValue, constants.xlSecondary)
            y2.HasMajorGridlines = bool(y2_grid) or False
            if y2_min not in (None, ""):
                y2.MinimumScale = y2_min
            if y2_max not in (None, ""):
                y2.MaximumScale = y2_max
            if y2_major not in (None, ""):
                y2.MajorUnit = y2_major
            if y2_format not in (None, ""):
                y2.TickLabels.NumberFormatLocal = y2_format
            if y2_title == False:
                y2.HasTitle = False
            elif y2_title not in (None, ""):
                y2.HasTitle = True
                y2.AxisTitle.Text = y2_title
                
        # 外枠の設定
        if frame_color == False:
            ch.ChartArea.Border.LineStyle = 0     # 枠なし
        elif frame_color not in (None, ""):
            ch.ChartArea.Format.Line.ForeColor.RGB = frame_color
            ch.ChartArea.Format.Line.Weight = p.get("frame_weight",0.75)  # 枠線の太さ(pt)
        else:
            ch.ChartArea.Format.Line.ForeColor.RGB = p.get("frame_color",RGB(217,217,217))  # 薄いグレー
            ch.ChartArea.Format.Line.Weight = p.get("frame_weight",0.75)                     # 枠線の太さ(pt)                  
        # ----------------------------------------------------------------------
    except Exception as e:
        print("フォーマットの設定でエラー:",e)
    try: 
        # 凡例を一度無効にする(例外あり)
        if legend == "right":
            ch.HasLegend = True
        elif legend not in (None, ""):
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
        elif legend is False:
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
        print("凡例かプロットエリアの調整でエラー:",e)
        
    return chart