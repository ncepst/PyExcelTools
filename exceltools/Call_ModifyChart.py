from exceltools import ModifyChart, RGB, add_shape, add_line, template
import xlwings as xw
import time

def main():
    # 計測開始
    t1 = time.time()

    wb = None
    try:
        wb = xw.books.active
        ws = wb.sheets.active
        wb.app.screen_updating = False

        import random
        r = random.randint(0, 100)

        chart = ModifyChart(title=f"タイトル{r}", width_cm=12, height_cm=8, title_space=0, x_major=15,
                            NS=1, plot_area_space="absolute", height_inc=0,width_inc=0)

        wb.app.screen_updating = True

    except Exception as e:
        print("Error:",e)

    finally:
        if wb is not None:
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
