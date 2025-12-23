#ExcelGraph.py
import xlwings as xw
from xlwings.constants import AxisType
from win32com.client import constants
import numpy as np
import time
import os

# 計測開始
t1 = time.time()

name = "グラフ名"
SeriesName = "系列名"
Title = name

# ------------------------------------------------------------
# path は適宜変更してください (Ctrl + Shift + C でパスのコピー)
# ------------------------------------------------------------
excel_path = r"C:\Users\*****\Desktop\test.xlsx"
Sheet = "Sheet1"

# RGBのヘルパー関数
def RGB(r, g, b):
    return r + g*256 + b*65536

# cm → pt 換算関数の定義 (1 point = 1/72 inch, 1 inch = 2.54 cm)
def cm_to_pt(cm):
    return cm * 72 / 2.54
    
# 指定のエクセルファイルがなければ作成する
if not os.path.exists(excel_path):
    wb = xw.Book()           # 新規Excelを開く
    wb.save(excel_path)
else:
    wb = xw.Book(excel_path) # 既存のExcelを開く
    
# 指定のワークシートがなければ作成する
if Sheet in [s.name for s in wb.sheets]:
    ws = wb.sheets[Sheet]
else:
    ws = wb.sheets.add(Sheet)

# エクセルの画面更新を無効にする
wb.app.screen_updating = False

# 上部に13行のセルを挿入する
ws.range("1:13").insert("down")

# デモデータ作成
data = 19
cell0 = np.linspace(-90, 90, data)
cell1 = np.cos(np.deg2rad(cell0))
cell = [cell0, cell1]

# データの貼り付け
ws.range(4,8).value = SeriesName
ws.range(2,8).value = Title
for n in range(2):
    for i in range(data):
        ws.range(n+3,i+9).value = cell[n][i]
        
# 数値の表示桁数の変更
for i in range(data):
    ws.range(3, i+9).number_format = "0"    # 小数0桁まで
for i in range(data):
    ws.range(4, i+9).number_format = "0.00" # 小数2桁まで
        
# cell1を数式に変更したい場合
for n in range(data):
    col = 9 + n                          # 9=I列
    col_letter = xw.utils.col_name(col)  # "I", "J", "K", ...
    ws.range(4, col).value = f"=cos({col_letter}3/180*pi())"
    
# ----------------------------------------------------------
# 散布図のエクセルグラフを作成する
# ----------------------------------------------------------
# (セル範囲入力) --------------------------------------------
start_range = "H3"
start = ws[start_range]
row = 2
col = data + 1
# (ターゲットセル計算) --------------------------------------
target_row = start.row + row - 1
target_col = xw.utils.col_name(start.column + col - 1)
target_range = f"{target_col}{target_row}"

# xlwings によるグラフ作成 ----------------------------------
chart = ws.charts.add(left=ws.range("A1").left+1,  # leftとtopは貼り付け位置の指定 (必須)
                      top=ws.range("A1").top+1,
                      width=cm_to_pt(12.54),       # widthとheightは大きさ指定 (省略可)
                      height=cm_to_pt(7.54)) 
chart.chart_type = 'xy_scatter_lines'
chart.set_source_data(ws.range(f'{start_range}:{target_range}'))
# ----------------------------------------------------------

# グラフのチャート名 (エクセル画面左上の表示で確認できる)
chart.name = name

# ChartObject(枠) → api[0]、Chart本体 → api[1]
ch = chart.api[1]

# グラフタイトル
ch.HasTitle = True
if ch.HasTitle:
    ch.ChartTitle.Text = Title
else:
    ch.ChartTitle.Text = ""

# 横軸のオプション
x_axis = ch.Axes(AxisType.xlCategory)
x_axis.MinimumScale = -90  # 最小値
x_axis.MaximumScale = 90   # 最大値  
x_axis.MajorUnit = 15      # 目盛間隔
x_axis.CrossesAt = -90     # 交差位置(縦軸との交点)
x_axis.TickLabels.NumberFormatLocal = "0"   # 小数0桁まで表示

# 横軸のタイトル
x_axis.HasTitle = True
if x_axis.HasTitle:
    x_axis.AxisTitle.Text = "角度 (deg.)"

# 縦軸のオプション
y_axis = ch.Axes(AxisType.xlValue)
y_axis.TickLabels.NumberFormatLocal = "0.0" # 小数1桁まで表示

# 縦軸のタイトル
y_axis.HasTitle = False
if y_axis.HasTitle:
    y_axis.AxisTitle.Text = ""

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
ch.HasLegend = False
    
# プロットエリアの調整
p = ch.PlotArea
p.InsideLeft   = p.InsideLeft
p.InsideTop    = p.InsideTop
p.InsideWidth  = p.InsideWidth
p.InsideHeight = p.InsideHeight+15  #下側に広げる

# 1つ目の系列の色を青にする
series = ch.SeriesCollection(1) 
series.Format.Line.ForeColor.RGB = RGB(68, 114, 196) # 線の色
series.MarkerForegroundColor = RGB(68, 114, 196)     # マーカー枠線の色
series.MarkerBackgroundColor = RGB(68, 114, 196)     # マーカー内部の色

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
    
# エクセルの画面更新を有効にする
wb.app.screen_updating = True

# エクセルファイルを保存する
wb.save()

# 計測終了
t2 = time.time()
elapsed_time = round(t2-t1,3)

print("処理時間:"+str(elapsed_time)+" s")
