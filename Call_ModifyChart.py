# Call_ModifyChart.py
import xlwings as xw
import time
from ModifyChart import ModifyChart, RGB

def main(mode=3):
    # 計測開始
    t1 = time.time()
    try:
        # 既に開いている Excel に接続
        wb = xw.books.active
        ws = wb.sheets.active
        wb.app.screen_updating = False
        # wb.app.calculation = 'manual'

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
                full_name = chart.Name                    # 例: "Sheet1 グラフ 1"
                space_index = full_name.find(" ")         # 空白の位置を取得
                # first_part = full_name[:space_index]
                second_part = full_name[space_index+1:]   # 空白を除いた残り
                print ("chart:",second_part)
                chart = ws.charts(second_part)
            else:
                print("アクティブなグラフはありません")
        
        import random
        r = random.randint(0, 100)
        ModifyChart(chart, title=f"タイトル{r}", width_cm=12, height_cm=8, title_space=0, x_major=1)

    except Exception as e:
        print("Error:",e)

    finally:
        # wb.app.calculate() 
        # wb.app.calculation = 'automatic'
        wb.app.screen_updating = True
    
    # 計測終了
    t2 = time.time()
    elapsed_time = round(t2-t1,3)
    print("処理時間:"+str(elapsed_time)+" s")
    
if __name__ == "__main__":
    main()
    
"""
#-----------------------------------------------------
    VBAマクロで Call_ModifyChart.py を実行するコード
(マクロのオプションから、ショートカットキーが設定できます)
#-----------------------------------------------------
Sub RunCallModifyChart()
    ' xlwings アドインがインストールされている必要があります
    RunPython ("import Call_ModifyChart; Call_ModifyChart.main()")
End Sub
"""
