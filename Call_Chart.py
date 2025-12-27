#Call_Chart.py
import xlwings as xw
from xlwings.constants import AxisType
from win32com.client import constants
import numpy as np
import time
import os

# 計測開始
t1 = time.time()

SeriesName = "系列名"
title = "タイトル"

# ------------------------------------------------------------
# path は適宜変更してください (Ctrl + Shift + C でパスのコピー)
# ------------------------------------------------------------
excel_path = r"C:\Users\*****\Desktop\test.xlsx"
Sheet = "Sheet1"
    
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
ws.range(2,8).value = title
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
    ws.range(5, col).value = f"=cos({col_letter}3/180*pi())*1.01"
    
from ScatterChart import ScatterChart, RGB

chart1 = ScatterChart(ws = ws,
                      start_range="H3",
                      row = 3,
                      col = data +1,
                      paste_range="A1",
                      width_cm=12.54, 
                      height_cm=7.54,
                      name = "",
                      title = "",
                      series_list = [{"color_RGB": (68,114,196)}],
                      x_title = "角度 (deg.)",
                      x_min = -90,
                      x_max = +90,
                      x_major = 15,
                      x_cross = -90,
                      x_format = "0",
                      y_format = "0.0",
                      legend="",
                      frame_color="", #黒=0
                      )  
    
# エクセルの画面更新を有効にする
wb.app.screen_updating = True

# エクセルファイルを保存する
wb.save()

# 計測終了
t2 = time.time()
elapsed_time = round(t2-t1,3)
print("処理時間:"+str(elapsed_time)+" s")