# create_sheet_index.py
import xlwings as xw

# 注意-------------------------------
# 既にある目次に戻るの行は削除されます。
# -----------------------------------

# RGBのヘルパー関数
def RGB(r, g, b):
    return r + g*256 + b*65536

def create_sheet_index():
    # pathは適宜変更してください-------------------------------
    wb = xw.Book(r"C:\Users\*****\Desktop\test.xlsx")
    # -------------------------------------------------------
    
    idx_name = "目次"

    # --- 既に「目次」シートがあれば削除 ---
    try:
        wb.sheets[idx_name].delete()
    except:
        pass

    # --- 新しい「目次」シートを先頭に作成 ---
    idx = wb.sheets.add(name=idx_name, before=wb.sheets[0])

    # タイトル
    idx.range("A1").value = "シート目次"
    idx.range("A1").api.Font.Bold = True
    idx.range("A1").api.Font.Size = 16

    r = 3  # 書き始め行

    # 各シートを目次に登録
    for ws in wb.sheets:
        if ws.name != idx_name:
            idx.api.Hyperlinks.Add(
                Anchor=idx.range(f"A{r}").api,
                Address="",
                SubAddress=f"'{ws.name}'!A1",
                TextToDisplay=ws.name
            )
            r += 1

    # 各シートに「目次へ戻る」リンクを作成
    for ws in wb.sheets:
        if ws.name == idx_name:
            continue

        # 既存の「目次へ戻る」リンクを探す
        found = False
        delete_row = None
        is_already_top = False

        for hl in ws.api.Hyperlinks:
            if hl.TextToDisplay == "目次へ戻る":
                delete_row = hl.Range.Row
                found = True
                if delete_row == 1:
                    is_already_top = True
                break

        # 既に1行目にリンクがある場合はスキップ
        if is_already_top:
            continue

        # 1行目以外にリンクがあった場合 → その行を削除
        if found and delete_row > 1:
            ws.api.Rows(delete_row).Delete()

        # 1行目に行を挿入
        ws.api.Rows(1).Insert()

        # A1 にリンク作成
        ws.api.Hyperlinks.Add(
            Anchor=ws.range("A1").api,
            Address="",
            SubAddress=f"'{idx_name}'!A1",
            TextToDisplay="目次へ戻る"
        )

        # 見やすいフォント設定
        ws.range("A1").api.Font.Bold = True
        ws.range("A1").api.Font.Size = 12
        ws.range("A1").api.Font.Color = RGB(0,0,255)

    import ctypes
    ctypes.windll.user32.MessageBoxW(0, "目次を作成しました", "完了", 0)

if __name__ == "__main__":

    create_sheet_index()
