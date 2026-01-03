## はじめに
本リポジトリの主要コードは、**Pythonで技術資料向けの Excel グラフの体裁を効率的に整える**ことを目的としています。  
主に **XYプロット（散布図）** を対象としており、測定データやシミュレーションデータなど、繰り返しグラフ作成が必要となる用途を想定しています。   

対象環境は Windows + Excel （Excel 2021 以降推奨）+ Python 3.10 以降 となります。  
使用する主なライブラリは `xlwings` と `win32com` です。

## 背景
Python から Excel グラフを作成する手段として、`openpyxl` や `xlwings` が利用できますが、    
以下のような課題を感じていました。
- Excel 上で手動作成したグラフと比べて、**書式が一致しないことが多い**
- 技術資料用途では、**細かな体裁調整が多数必要**になる
- `matplotlib` で作成した画像形式のグラフを貼り付ける方法では、  
  Excel COM を介さないため処理速度の面でメリットはあるものの、  
  Excel 上での再編集ができず、用途に合わないケースがある
- VBAによる Excel グラフ制御の情報は比較的多い一方で、  
  **Pythonを用いて書式設定まで自動化した事例は少ないように思われる**
    
また、Excel グラフの体裁を細かく調整するには多数のパラメータ指定が必要となり、  
手軽に自動化プログラムへ組み込むことが難しいという問題もありました。

そこで本リポジトリでは、  
**グラフ作成だけでなく書式設定までを含めて自動化プログラムに組み込めること**を目的とし、  
繰り返し利用可能な汎用関数としてコードを作成しました。  

これにより、技術資料作成や解析用途における作業効率の向上を図っています。  

## 概要 

### 記事とサンプルコード

- Qiita に投稿した記事: [Python xlwings でエクセルグラフの体裁調整を自動化](https://qiita.com/ncepst/items/54c63974242bb9e18c23)
- サンプルコード: [excel_graph_sample.py](https://github.com/ncepst/PyExcelTools/blob/main/excel_graph_sample.py)  

上記記事を発展させ、グラフ作成とその書式設定を関数化したコードとして整理しています。  
グラフ書式設定パラメータは`PRESET`で管理しています。

### 構成ファイル

**グラフ作成**
- [ScatterChart.py](https://github.com/ncepst/PyExcelTools/blob/main/ScatterChart.py) — グラフ作成関数  
- [Call_ScatterChart.py](https://github.com/ncepst/PyExcelTools/blob/main/Call_ScatterChart.py) — ScatterChart 関数の呼び出し例

**既存グラフの編集**

- [ModifyChart.py](https://github.com/ncepst/PyExcelTools/blob/main/ModifyChart.py) — グラフ変更関数  
- [Call_ModifyChart.py](https://github.com/ncepst/PyExcelTools/blob/main/Call_ModifyChart.py) — ModifyChart 関数の呼び出し例

**VBA連携**

- [PythonFromVBA.bas](https://github.com/ncepst/PyExcelTools/blob/main/PythonFromVBA.bas)

Pythonスクリプトは Excelマクロから実行することもでき、  
マクロにショートカットキーを割り当てると便利です。  

付属のVBAコードには、  
● 自動判定の範囲 または 選択範囲 でグラフを作成するマクロ  
● `Call_ModifyChart.py` を呼び出して選択中のグラフの体裁編集ができるマクロ    
が含まれており、柔軟な運用が可能です。

---

## グラフ書式設定の指定方法

グラフの書式設定の項目のうち、用途ごとに指定が必要な項目が関数の引数としてリスト化されています。  
`ScatterChart.py`は必須引数が`ws, start_range, paste_range`の3つ、  
`ModifyChart.py`は必須引数が`chart`の1つのみで、  
残りの54個の任意引数により、設定項目を指定します。

また、`PRESET`は **dict形式** で関数外に定義されており、  
外側の辞書のキー`"excel2021"`に、Excel 2021 の散布図の書式設定が格納されています。  
各グラフ設定パラメータはネスト(入れ子)辞書として管理されており、必要に応じて上書き・変更が可能です。
```python
"std": {
        "title_font_color": RGB(0,0,0),
        "axis_title_font_color":RGB(0,0,0), 
        "axis_tick_font_color": RGB(0,0,0),
    }
PRESET["std"] = {**PRESET["excel2021"], **PRESET["std"]}
```
上記のようにベースとなる設定を維持したまま一部を上書きすることで、  
詳細な書式設定を簡単に保存できます。  
関数の引数で`preset`を指定することができ、デフォルト引数では`preset="std"`が適用されます。

呼び出し側でも、PRESETの更新が可能で、  
以下では、グリッド線の表示設定を一時的に False に変更しています。  
呼び出し側での変更は その Python セッション内で (スクリプトが終了するまで) 有効となります。
```python
from ScatterChart import ScatterChart, PRESET
PRESET["std"]["x_major_grid"] = False
chart1 = ScatterChart(....)

PRESET["std"]["x_major_grid"] = True
chart2 = ScatterChart(....)
```

任意引数のうち、系列ごとの書式設定は `series_list = [{"name":"系列1"},{"name":"系列2"}, ...]`  
で指定します。各系列は **dict形式** で定義し、系列をまとめてリスト`series_list`に格納します。

`ModifyChart.py`を`series_list = None` で実行する場合には、  
処理する系列数を関数の引数`NS`で指定してください。  
デフォルト引数では`NS=1`となっています。  
また、必須引数である`chart`を得るためのコードは`Call_ModifyChart.py`に記載しています。

一方、`ScatterChart.py`では、`row` と `col` の値から`NS`が自動計算されます。  
`row` や `col` が `None` の場合は、`start_range`の入力値から`.end('down').row` 等を使って自動取得されます。   
`paste_range`ではグラフを貼り付ける左上の座標として、どのセルの左上を基準にするのかを指定します。

### ModifyChart関数の引数説明
```python
def ModifyChart(chart,                        # ExcelのChartオブジェクト
                ws = None,                    # None (系列ごとにRange指定する場合のみWorksheetオブジェクトを指定)
                preset = "std",               # プリセットスタイル名
                width_cm = None,              # グラフ幅(cm), Noneで変更なし
                height_cm = None,             # グラフ高さ(cm), Noneで変更なし
                name = None,                  # グラフ名変更, Noneで変更なし
                title: str|bool|None = None,  # タイトル文字列, Falseで無効化, Noneで変更なし
                title_font_color = None,      # タイトルフォント色(RGB)
                title_font_size = None,       # タイトルフォントサイズ
                title_space = +0,             # タイトルとグラフの間隔(pt)
                NS = 1,                       # データ系列数
                series_list: list[dict]|None = None,  # 各系列の設定
                style = None,                 # 線＋マーカーのスタイル
                smooth = None,                # Trueで曲線、Falseで折れ線
                marker = None,                # マーカー種類: "C":●, "S":■, "D":◆, "T":▲, "N":なし
                alpha = None,                 # 線の透明度(0~1)
                x_title: str|bool|None = None,# X軸タイトル文字列。Falseで無効化、Noneで変更なし
                x_title_space = +0,           # プロットエリアを下側に広げる場合はマイナス
                x_min: float|str|None = None, # X軸最小値, "auto"で自動調整, Noneで変更なし
                x_max: float|str|None = None, # X軸最大値, "auto"で自動調整, Noneで変更なし
                x_major: float|None = None,   # X軸主目盛間隔, Noneで変更なし, グリッド線表示設定はPRESET
                x_minor: float|None = None,   # X軸副目盛間隔, Noneで変更なし, グリッド線表示設定はPRESET
                x_cross = None,               # Y軸との交差位置, Noneで変更なし
                x_format = None,              # X軸表示形式 ("0.00", "0.0E+00", "0%" など), Noneで変更なし
                x_log: bool|None = None,      # Trueで対数表示, Noneで変更なし
                y_title: str|bool|None = None,# Y軸タイトル, Noneで変更なし
                y_title_space = +0,           
                y_min: float|str|None = None,  
                y_max: float|str|None = None,  
                y_major: float|None = None,    
                y_minor: float|None = None,    
                y_cross = None,                    # X軸との交差位置, Noneで変更なし        
                y_format = None,              
                y_log: bool|None = None,       
                y2_title: str|bool|None = None,    # Y軸タイトル文字列, Falseで無効化, Noneで変更なし
                y2_min: float|str|None = None,
                y2_max: float|str|None = None,
                y2_major: float|None = None,
                y2_minor: float|None = None,
                y2_format = None,
                y2_log: bool|None = None,       
                frame_color: bool|int|None = None, # グラフ枠色 (False:枠なし, 0:黒枠, Noneでpreset)
                legend: bool|str|None = None,      # 凡例表示(False:非表示, Noneで変更なし, "right"で右側に表示) 
                legend_font_size = None,           # 凡例のフォントサイズ, Noneで変更なし
                legend_width_inc = 0,              # 凡例ボックスの幅増減(pt)
                legend_height_inc = 0,             # 凡例ボックスの高さ増減(pt)
                legend_right_space = 0,            # 凡例の右端 = プロットエリアの右端を基準とした凡例位置制御
                legent_top_space = 0,              # 凡例の上端 = プロットエリアの上端を基準とした凡例位置制御
                transparent_bg: bool|None = None,  # 背景を透明化する場合はTrue
                chart_type = None,                 # "bar"で棒グラフに変更
                x_bold_line: float|None = None,    # x_bold_line=0でx=0が太線
                y_bold_line: float|None = None,    # y_bold_line=0でy=0が太線
                plot_area_space: str = "relative", # プロットエリア調整時の基準位置を"absolute"/"relative"で指定
                width_inc = 0,                     # プロットエリアの幅増減(pt)
                height_inc = 0,                    # プロットエリアの高さ増減(pt)
                ):
```
- 型ヒントには `Python 3.10 以降`の記法を使用しています。 `ModifyChart.py`のみ型ヒントを使用しました。
  
- `smooth`, `alpha`, `line_weight`, `marker` については、series cfg と PRESET にも同様の項目があり、  
優先順位としては、`series cfg > 引数 > preset` となります。  
series cfg (設定なしで自動的にNone) と 引数が None で preset の設定が適用されます。
 
- `ScatterChart`関数の場合には、上記の引数: `NS`, `chart`の代わりに、  
必須引数:`start_range`, `paste_range`、 任意引数:`row` ,`col` が加わって、`ws`は必須引数に変わります。  
戻り値はChartオブジェクトとなります。

- `chart_type = "bar"` を指定すると、集合縦棒グラフ (`xlColumnClustered`) が表示されます。   
  ただしこの場合、内部処理でtry構文の例外が発生し、軸の設定・軸のフォント設定には未対応です。
  
- 引数で `chart_type = "line"`を指定し、そのうち一部の系列の chart_type を "bar" にすると、  
  棒グラフ+折れ線グラフの組み合わせグラフ を表示できます。

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
        "trendline_name":None,        # 近似曲線の名前
        "trendline_color":None,       # 近似曲線の色
        "trendline_weight":None,      # 近似曲線の太さ
        "trendline_style":None,       # 近似曲線のDashstyle(solid,dash)
        "trendline_option": None,     # 近似曲線オプション("eq+r2")
        "legend":None,                # 系列を凡例に入れるか選択
        "data_label":None             # データラベルの表示
    },
]
```
### サブ関数
`ModifyChart.py`には、`add_shape`, `add_line` の2つの関数が、メイン関数`ModifyChart` とは独立に定義されており、  
`from ModifyChart import MofifyChart, RGB, add_shape, add_line` のようにインポートして使います。  
`add_shape` は、引数に指定する`chart`のプロットエリア内に、帯状の長方形や注釈用のテキストボックスを追加します。  
`add_line` は、引数に指定する`chart`のプロットエリア内に、線もしくは破線を追加します。  

詳細な使い方は、`ModifyChart.py` の各関数定義の下にある docstring をご参照ください。

### PRESET["excel2021"] 一覧
```python
PRESET = {
    # "Aptos Narrow 本文"は Excel 2021以降のみ
    "excel2021": {
        "title_font_size": 14,
        "title_font_color": RGB(89,89,89),
        "title_font_bold": False,
        "title_font_name": "Aptos Narrow 本文",
        "axis_title_font_size": 10,
        "axis_title_font_name": "Aptos Narrow 本文",
        "axis_title_font_color": RGB(89,89,89),
        "axis_title_font_bold": False,
        "axis_tick_font_color": RGB(89, 89, 89),
        "axis_tick_font_size": 9,
        "axis_tick_font_name": "Aptos Narrow 本文",
        "axis_line": True,
        "axis_line_color": RGB(191, 191, 191),
        "axis_line_weight": 0.75,    
        "x_major_grid": True,
        "y_major_grid": True,
        "x_minor_grid": None,
        "y_minor_grid": None,
        "major_grid_color": RGB(217, 217, 217),
        "major_grid_weight": 0.75,
        # TickMark: None, Inside, Outside, Cross
        "x_major_tickmark": constants.xlTickMarkNone,  # 目盛の内向き/外向きなし
        "y_major_tickmark": constants.xlTickMarkNone,
        "x_minor_tickmark": None,
        "y_minor_tickmark": None,
        # 外枠の有無は引数で指定してください
        "frame_color": RGB(217,217,217),
        "frame_weight": 0.75,
        "plot_area_frame": None,
        "plot_area_frame_color": RGB(0,0,0),
        "plot_area_frame_weight": 1.0,
        "style": "line+marker",
        "smooth": True,
        "alpha": None,
        "line_weight": 1.5,
        "marker": "C",
        "markersize": 5,
        "y2_major_grid": False,  # 副軸グリッド表示なし
        "y2_minor_grid": False,
        "y2_major_tickmark": constants.xlTickMarkNone,
        "y2_minor_tickmark": None,
    },
}
```
### 補足
- 透明度(0～1)の`alpha`は`matplotlib`に由来です。  
- 線の太さ`line_weight`はExcel COM のプロパティ名`Line.Weight`に由来します。 
  (matplotlibでは`linewidth`)  
- マーカーの種類(`"C","S","D","T","N"`)は頭文字に由来します。
  (matplotlibの`"o","s","D","^",""`は使えません)  
- `style`では、`matplotlib`由来の表記で`"--"`や`"-."`が利用できますが、  
破線の種類は上記の2種類のみ対応しています。  
- 関数の引数が多いので、同じ引数を二度渡さないように注意して下さい。  
  (二重指定すると `SyntaxError` になります)
- 系列の色指定については`RGB(r,g,b)`の他に以下の文字列でも指定できます。
```python
COLOR_NAME_TO_RGB = {
              "blue": RGB(68,114,196),
              "orange":RGB(237,125,49),
              "green":RGB(112,173,71),
              "yellow":RGB(255,192,0),
              "purple":RGB(112,48,160),
              "brown":RGB(192,0,0),
              "navy":RGB(0,32,96),
              "gray":RGB(165,165,165),
              "teal":RGB(0,128,128),
              "cyan": RGB(0,255,255),
              "magenta": RGB(255,0,255),
              }
```
## License
MIT License  
Copyright (c) 2025 ncepst
