## 概要

**Pythonで技術資料向けのエクセルグラフの体裁を素早く整える**ことを目的としています。  
対象環境は Windows + Excel （Excel 2021 以降推奨）となります。 

参考記事とサンプルコード

- Qiitaに記事を投稿しています: [Python xlwings でエクセルグラフの体裁調整を自動化](https://qiita.com/ncepst/items/54c63974242bb9e18c23)
- サンプルコード: [excel_graph_sample.py](https://github.com/ncepst/PyExcelTools/blob/main/excel_graph_sample.py)  

記事を発展させて、グラフ作成を関数化したコードを作成しています。  
`PRESET`にはグラフ書式設定パラメータが格納されています。

- [ScatterChart.py](https://github.com/ncepst/PyExcelTools/blob/main/ScatterChart.py) — グラフ作成関数  
- [Call_ScatterChart.py](https://github.com/ncepst/PyExcelTools/blob/main/Call_ScatterChart.py) — ScatterChart 関数の呼び出し例

さらに、既存グラフを変更する関数を作成しました。

- [ModifyChart.py](https://github.com/ncepst/PyExcelTools/blob/main/ModifyChart.py) — グラフ変更関数  
- [Call_ModifyChart.py](https://github.com/ncepst/PyExcelTools/blob/main/Call_ModifyChart.py) — ModifyChart 関数の呼び出し例

Pythonスクリプトは Excelマクロから実行することもでき、マクロにショートカットキーを割り当てると便利です。   
付属のVBAコードには、自動判定の範囲 または 選択範囲 でグラフを作成するマクロと、  
`Call_ModifyChart.py` を呼び出して選択中のグラフの体裁編集ができるマクロが含まれており、  
柔軟な運用が可能です。

- [PythonFromVBA.bas](https://github.com/ncepst/PyExcelTools/blob/main/PythonFromVBA.bas)

---

## グラフ書式設定の指定方法

グラフの書式設定の項目は関数の引数としてリスト化されています。  
`ScatterChart.py`は必須引数が`ws, start_range, paste_range`の3つ、  
`ModifyChart.py`は必須引数が`chart`の1つのみで、  
残りの53個の任意引数により、設定項目を指定します。

また、`PRESET`は **dict形式** で関数外に定義されており、  
外側の辞書のキー`"excel2021"`に、Excel 2021 の散布図の書式設定が格納されています。  
各グラフ設定パラメータはネスト(入れ子)辞書として管理されており、必要に応じて上書き・変更が可能です。
```python
"std": {
        "axis_title_font_color":RGB(0,0,0), 
        "axis_tick_font_color": RGB(0,0,0),
    }
PRESET["std"] = {**PRESET["excel2021"], **PRESET["std"]}
```
上記のようにベースとなる設定を維持したまま一部を上書きすることで、  
詳細な書式設定を簡単に保存できます。  
関数のデフォルト引数としては、`preset="std"`が適用されます。

任意引数のうち、系列ごとの書式設定は `series_list = [{"name":"系列1"},{"name":"系列2"}, ...]`  
で指定します。各系列は **dict形式** で定義し、複数系列の場合はリストとしてまとめます。

`ModifyChart.py`を`series_list = None` で実行する場合には、  
処理する系列数を関数の引数`NS`で指定してください。  
デフォルト引数では`NS=1`となっています。  
必須引数である`chart`を得るためのコードは`Call_ModifyChart.py`に記載しています。

一方、`ScatterChart.py`では、`row` と `col` の値から`NS`が自動計算されます。  
`row` や `col` が `None` の場合は、`start_range`の入力値から`.end('down').row` 等を使って自動取得されます。   
`paste_range`ではグラフを貼り付ける左上の座標として、どのセルの左上を基準にするのかを指定します。  

### series_listで指定可能なkeyとそのデフォルト値
以下から任意のkeyのみ設定可能です。
```python
series_list = [
    {
        "name": "系列1",               # 系列名
        "color": RGB(68,114,196),     # 色
        "style": "line+marker",       # スタイル
        "marker": "C",                # マーカー: "C":●, "S":■, "D":◆, "T":▲
        "size": 5,                    # マーカーサイズ
        "weight": 1.5,                # 線の幅 (pt)
        "smooth": True,               # 曲線(True) or 折れ線(False)
        "alpha": None,                # 線の透明度(0~1), デフォルト None
        "axis": "primary",            # "y2"で副軸
        "XValues":None,               # 系列のXの値(ExcelのRange指定)
        "Values":None,                # 系列のYの値(ExcelのRange指定)
        "sheet":None,                 # 系列のデータがあるシート名
        "chart_type": None,           # グラフ種類の変更(デフォルトは散布部)
        "trendline": None,            # 近似曲線
        "trendline_option": None,     # 近似曲線オプション("eq+r2")
        "legend":None,                # 系列を凡例に入れるか選択
        "data_label":None             # データラベルの表示
    },
]
```
### サブ関数
`ModifyChart.py`には、`add_shape`, `add_line` の2つの関数が`ModifyChart` とは独立に定義されており、  
`from ModifyChart import MofifyChart, RGB, add_shape, add_line` でインポートして使います。  
`add_shape` は、引数に指定する`chart`のプロットエリア内に、帯状の長方形や注釈用のテキストボックスを追加します。  
`add_line` は、引数に指定する`chart`のプロットエリア内に、線もしくは破線を追加します。  

詳細な使い方は、`ModifyChart.py` の各関数定義の下にある docstring をご参照ください。

## License
Copyright (c) 2025 ncepst
