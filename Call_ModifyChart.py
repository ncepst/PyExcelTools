#Call_ModifyChart
import xlwings as xw
import time
from ModifyChart import ModifyChart

# 計測開始
t1 = time.time()

mode = 3

# 既に開いている Excel に接続
wb = xw.books.active
ws = wb.sheets.active

# 既存グラフをインデックスで指定する (0,1,2..)
if mode == 1:
    chart = ws.charts[0]

# 既存グラフをグラフ名で指定する
elif mode == 2:
    chart = ws.charts("グラフ 1")

# 選択中のグラフを指定する
elif mode == 3:
    app = xw.apps.active
    chart = app.api.ActiveChart
    if chart is not None:
        full_name = chart.Name                      # 例: "Sheet1 グラフ 1"
        # 空白の位置を取得
        space_index = full_name.find(" ")
        first_part = full_name[:space_index]        # 先頭から最初の空白まで
        second_part = full_name[space_index+1:]     # 空白を除いた残り
        print ("chart:",second_part)
        chart = ws.charts(second_part)
    else:
        print("アクティブなグラフはありません")

ModifyChart(chart,Title="測定結果",width_cm=12,height_cm=8, Title_space=+20)

# 計測終了
t2 = time.time()
elapsed_time = round(t2-t1,3)
print("処理時間:"+str(elapsed_time)+" s")