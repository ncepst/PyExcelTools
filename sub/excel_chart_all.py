import xlwings as xw
import os

app = xw.App(visible=False)     # Excel を表示しない
wb = app.books.open(r"C:\Users\*****\Desktop\test.xlsx")

def RGB(r, g, b):
    return r + (g << 8) + (b << 16)

ws = wb.sheets[0]               # 1つ目のシートを指定

# -----------------------------------------------------
# 1) シート内のグラフ名を全部表示して確認
# -----------------------------------------------------
chart_count = ws.api.ChartObjects().Count
print("Chart count:", chart_count)

for i in range(1, chart_count + 1):
    obj = ws.api.ChartObjects(i)
    print(i, obj.Name)

# -----------------------------------------------------
# 2) たとえば 1番目のグラフを取得（必要なら Name 指定も可）
# -----------------------------------------------------
obj = ws.api.ChartObjects(2)
chart = obj.Chart   # Excel の本物の Chart オブジェクト

# --- 系列の色変更（RGB） ---
series = chart.SeriesCollection(1)     # 1つ目の系列
series.Format.Line.ForeColor.RGB = RGB(0, 255, 0) # 線の色
series.MarkerForegroundColor = RGB(0, 255, 0)     # マーカー枠線の色
series.MarkerBackgroundColor = RGB(0, 255, 0)     # マーカー内部の色

wb.save()

wb2 = app.books.add()

# COM オブジェクトを取得して、コピー
ws1 = wb.sheets[0].api  
target = wb2.sheets[0].api
ws1.Copy(Before=target)

# シート名を変更
ws2 = wb2.sheets[0]
ws2.name = "Copyファイル"

# コピー元の保存先フォルダを取得
folder = os.path.dirname(wb.fullname)

wb2.save(os.path.join(folder, "test_v2.xlsx"))
wb2.close()
app.quit()