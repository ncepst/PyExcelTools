#Call_ModifyChart
import xlwings as xw
import time
from ModifyChart import ModifyChart

# 計測開始
t1 = time.time()

# 既に開いている Excel に接続
wb = xw.books.active
ws = wb.sheets.active

# 既存グラフを順番で指定する場合（0, 1, 2...）
# chart = ws.charts[0]

# 既存グラフをグラフ名で指定する場合
chart = ws.charts("グラフ 1")

ModifyChart(chart,Title="測定結果",width_cm=12,height_cm=8, Title_space=+20)

# 計測終了
t2 = time.time()
elapsed_time = round(t2-t1,3)
print("処理時間:"+str(elapsed_time)+" s")