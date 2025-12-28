# Excel グラフ作成・編集ツール

Excel 2021 の標準スタイルに寄せて、**研究資料・技術資料向けのグラフの体裁を素早く整える**ことを目的としています。  
対象環境は Windows + Excel 2021 以降 です。 

サンプルコードと参考記事

- xlwings によるエクセルグラフ作成自動化の記事を投稿しています。: [Qiita 記事](https://qiita.com/Bencepst/items/54c63974242bb9e18c23)
- サンプルコード: [excel_graph_sample.py](https://github.com/ncepst/PyExcelTools/blob/main/excel_graph_sample.py)  

記事を元に、グラフ作成を関数化したコードを作成しています。

- [ScatterChart.py](https://github.com/ncepst/PyExcelTools/blob/main/ScatterChart.py) — グラフ作成関数  
- [Call_ScatterChart.py](https://github.com/ncepst/PyExcelTools/blob/main/Call_ScatterChart.py) — ScatterChart 関数の呼び出し例

さらに、既存グラフを変更する関数を作成しました。

- [ModifyChart.py](https://github.com/ncepst/PyExcelTools/blob/main/ModifyChart.py) — グラフ変更関数  
- [Call_ModifyChart.py](https://github.com/ncepst/PyExcelTools/blob/main/Call_ModifyChart.py) — ModifyChart 関数の呼び出し例

Pythonは Excelマクロから実行することもでき、マクロにショートカットキーを割り当てると便利です。   
付属のVBAコードには、自動判定の範囲 または 選択範囲 でグラフを作成するマクロと、  
Call_ModifyChart.py を呼び出して選択中のグラフの体裁編集ができるマクロが含まれており、柔軟な運用が可能です。

- [PythonFromVBA.bas](https://github.com/ncepst/PyExcelTools/blob/main/PythonFromVBA.bas)

---

## グラフ書式設定の指定方法

グラフの書式設定は関数の引数としてリスト化されています。  
そのうち、系列ごとの書式設定は **dict形式** で指定可能です。

- 指定可能なキーは任意で設定可能
- `series_list = None` の場合は、処理する系列数を関数の引数 `NS` で指定

### series_list の例

```python
series_list = [
    {
        "name": "系列1",               # 系列名
        "color": RGB(68,114,196),     # 色
        "style": "line+marker",       # スタイル
        "marker": "C",                # マーカー: C, S, D, T
        "size": 5,                    # マーカーサイズ
        "weight": 1.5,                # 線の幅 (pt)
        "smooth": True,               # 曲線(True) or 折れ線(False)
        "alpha": None,                # 線の透明度(0~1), デフォルト None
        "axis": "primary",            # "y2"で副軸
        "chart_type": None,           # デフォルトは散布図
        "trendline": None,            # 近似曲線
        "trendline_option": None,     # 近似曲線オプション("eq+r2")
        "legend":None,                # 系列を凡例に入れるか選択
        "data_label":None             # データラベルの表示
    },
]
