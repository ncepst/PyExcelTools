## 概要

**Pythonで技術資料向けのエクセルグラフの体裁を素早く整える**ことを目的としています。  
対象環境は Windows + Excel 2021 以降 です。 

参考記事とサンプルコード

- xlwings によるエクセルグラフ作成自動化の記事を投稿しています。: [Qiita 記事](https://qiita.com/Bencepst/items/54c63974242bb9e18c23)
- サンプルコード: [excel_graph_sample.py](https://github.com/ncepst/PyExcelTools/blob/main/excel_graph_sample.py)  

記事を発展させて、グラフ作成を関数化したコードを作成しています。  
PRESETで、詳細な書式設定を一覧化しています。

- [ScatterChart.py](https://github.com/ncepst/PyExcelTools/blob/main/ScatterChart.py) — グラフ作成関数  
- [Call_ScatterChart.py](https://github.com/ncepst/PyExcelTools/blob/main/Call_ScatterChart.py) — ScatterChart 関数の呼び出し例

さらに、既存グラフを変更する関数を作成しました。

- [ModifyChart.py](https://github.com/ncepst/PyExcelTools/blob/main/ModifyChart.py) — グラフ変更関数  
- [Call_ModifyChart.py](https://github.com/ncepst/PyExcelTools/blob/main/Call_ModifyChart.py) — ModifyChart 関数の呼び出し例

Pythonは Excelマクロから実行することもでき、マクロにショートカットキーを割り当てると便利です。   
付属のVBAコードには、自動判定の範囲 または 選択範囲 でグラフを作成するマクロと、  
Call_ModifyChart.py を呼び出して選択中のグラフの体裁編集ができるマクロが含まれており、  
柔軟な運用が可能です。

- [PythonFromVBA.bas](https://github.com/ncepst/PyExcelTools/blob/main/PythonFromVBA.bas)

---

## グラフ書式設定の指定方法

グラフの書式設定の項目は関数の引数としてリスト化されています。  
ScatterChart.pyは必須引数が`ws, start_range, paste_range`の3つ、  
ModifyChart.pyは必須引数が`chart`の1つのみで、  
グラフの書式設定に関する任意引数が53個あります。

引数のうち、系列ごとの書式設定 `series_list = [{"name":"系列1"},{"name":"系列2"}]`については、  
**dict形式** で指定します。

`series_list = None` でModifyChart.pyを実行する場合には、  
処理する系列数を関数の引数`NS`で指定してください。  
デフォルト引数では`NS=1`となっています。  
ScatterChart.pyでは、row, colの値から`NS`が自動計算されます。  

### series_listで指定可能なkeyとそのデフォルト値
以下から任意のkeyのみ設定可能です。
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
