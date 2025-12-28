PyExcelTools: Excel グラフ作成・編集ツール

Excel 2021 の標準スタイルに寄せて、研究資料・技術資料向けのグラフの体裁を素早く整えることを目的としています。
対象環境は Windows + Excel 2021 以降 です。

1. サンプルコードと参考記事

- サンプルコード: excel_graph_sample.py 
- xlwings によるエクセルグラフ作成自動化の記事: https://qiita.com/Bencepst/items/54c63974242bb9e18c23

記事を元に、グラフ作成を関数化したコードを作成しています。

- [ScatterChart.py](https://github.com/ncepst/PyExcelTools/blob/main/ScatterChart.py) — グラフ作成関数  
- [Call_ScatterChart.py](https://github.com/ncepst/PyExcelTools/blob/main/Call_ScatterChart.py) — ScatterChart 関数の呼び出し例

さらに、既存グラフを変更する関数も作成しました。

- [ModifyChart.py](https://github.com/ncepst/PyExcelTools/blob/main/ModifyChart.py) — グラフ変更関数  
- [Call_ModifyChart.py](https://github.com/ncepst/PyExcelTools/blob/main/Call_ModifyChart.py) — ModifyChart 関数の呼び出し例

Call_ModifyChart.pyを呼び出すExcelマクロで、ショートカットキーを割り当てると便利です。

- [PythonFromVBA.bas](https://github.com/ncepst/PyExcelTools/blob/main/PythonFromVBA.bas)

---

グラフの書式設定の項目は関数の引数としてリスト化されています。
そのうち 系列ごとの書式設定はdict形式で指定します。
指定可能なkeyと、そのデフォルト値は以下の通りです。
任意のkeyのみ設定可能です。
series_list = None のまま ModifyChart.pyを実行する場合には、
処理する系列数を関数の引数NSで指定してください。

series_list = [
    {
    "name":"系列1",                 # 系列名
    "color": RGB(68,114,196),       # 色
    "style":"line+marker",          # スタイル
    "marker":"C",                   # マーカー: C, S, D, T
    "size":5,                       # マーカーサイズ
    "weight":1.5,                   # 線の幅 (pt)
    "smooth":True,                  # 曲線 or 折れ線か
    "alpha":None,                   # 線の透明度(0~1), デフォルト0
    "axis":"primary",               # "y2"で副軸
    "chart_type":None,              # デフォルト散布図
    "trendline":None,               # 近似曲線
    "trendline_option":None         # 近似曲線オプション("eq+r2")
    }
    ],
