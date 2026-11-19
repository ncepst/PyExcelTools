"""複数の埋め込みグラフに ModifyChart の設定を適用します。"""
from copy import deepcopy
from inspect import signature
from pathlib import Path
from typing import Any, Literal

from .ModifyChart import ModifyChart


def BatchModifyCharts(
    target: Literal["sheet", "book", "selected", "dir"] = "sheet",
    *,
    ws: Any = None,
    wb: Any = None,
    directory: str | Path | None = None,
    recursive: bool = False,
    preset: Literal["excel2021", "std", "no_grid", "grid", "no_change"] = "std",   # プリセットスタイル名
    width_cm = None,              # グラフ幅(cm), Noneで変更なし
    height_cm = None,             # グラフ高さ(cm), Noneで変更なし
    name = None,                  # グラフ名変更, Noneで変更なし
    title: Literal[None, False, "タイトル"] = None,  # タイトル文字列, Falseで無効化, Noneで変更なし
    title_font_color = None,      # タイトルフォント色(RGB)
    title_font_size = None,       # タイトルフォントサイズ
    title_space = +0,             # タイトルとグラフの間隔(pt)
    NS: int | None = None,                  # データ系列数。Noneで各グラフから自動取得
    series_list: list[dict]|None = None,  # 各系列の設定
    style: Literal[None, "line+marker", "line", "marker", "dash", "dash+marker", "chain"] = None, # 線＋マーカーのスタイル
    smooth: bool = None,         # Trueで曲線、Falseで折れ線
    marker: Literal["C", "S", "D", "T", "N"] = None,  # マーカー種類: "C":●, "S":■, "D":◆, "T":▲, "N":なし
    alpha = None,                 # 線の透明度(0~1)
    x_title: Literal[None, False, "X軸タイトル"] = None,  # X軸タイトル文字列。Falseで無効化、Noneで変更なし
    x_title_space = +0,           # プロットエリアを下側に広げる場合はマイナス
    x_min: float|str|None = None, # X軸最小値, "auto"で自動調整, Noneで変更なし
    x_max: float|str|None = None, # X軸最大値, "auto"で自動調整, Noneで変更なし
    x_major: float|None = None,   # X軸主目盛間隔, Noneで変更なし, グリッド線表示設定はPRESET
    x_minor: float|None = None,   # X軸副目盛間隔, Noneで変更なし, グリッド線表示設定はPRESET
    x_cross = None,               # Y軸との交差位置, Noneで変更なし
    x_format: Literal[None, "0", "0.0", "0.00", "0.0E+00", "0%", "#,##0"] = None,          # X軸表示形式 ("0.00", "0.0E+00", "0%" など), Noneで変更なし
    x_log: Literal[None, True] = None,      # Trueで対数表示, Noneで変更なし
    y_title: Literal[None, False, "Y軸タイトル"] = None, # Y軸タイトル, Noneで変更なし
    y_title_space = +0,
    y_min: float|str|None = None,
    y_max: float|str|None = None,
    y_major: float|None = None,
    y_minor: float|None = None,
    y_cross = None,                    # X軸との交差位置, Noneで変更なし
    y_format: Literal[None, "0", "0.0", "0.00", "0.0E+00", "0%", "#,##0"] = None,
    y_log: Literal[None, True] = None,
    y2_title: Literal[None, False, "副軸タイトル"]  = None, # Y軸タイトル文字列, Falseで無効化, Noneで変更なし
    y2_min: float|str|None = None,
    y2_max: float|str|None = None,
    y2_major: float|None = None,
    y2_minor: float|None = None,
    y2_format: Literal[None, "0", "0.0", "0.00", "0.0E+00", "0%", "#,##0"] = None,
    y2_log: Literal[None, True] = None,
    frame_color: Literal[None, False, 0] = None, # グラフ枠色 (False:枠なし, 0:黒枠, Noneでpreset)
    legend: Literal[False, None, "right", "TR+bw+fb", "TC+bw+fb", "TL+bw+fb", "BL+bw+fb", "BC+bw+fb", "BR+bw+fb", "TR", "TL", ""] = None,    # 凡例表示(False:非表示, Noneで変更なし, "right"で右側に表示)
    legend_font_size = None,           # 凡例のフォントサイズ, Noneで変更なし
    legend_width_inc = 0,              # 凡例ボックスの幅増減(pt)
    legend_height_inc = 0,             # 凡例ボックスの高さ増減(pt)
    legend_right_space = 0,            # 凡例の右端 = プロットエリアの右端を基準とした凡例位置制御
    legent_top_space = 0,              # 凡例の上端 = プロットエリアの上端を基準とした凡例位置制御
    transparent_bg: bool|None = None,  # 背景を透明化する場合はTrue
    chart_type: Literal[None, "bar"] = None,  # "bar"で棒グラフに変更
    x_bold_line: float|None = None,    # x_bold_line=0でx=0が太線
    y_bold_line: float|None = None,    # y_bold_line=0でy=0が太線
    plot_area_space: Literal["relative", "absolute"] = "relative", # プロットエリア調整時の基準位置を"absolute"/"relative"で指定
    width_inc = 0,                     # プロットエリアの幅増減(pt)
    height_inc = 0,                    # プロットエリアの高さ増減(pt)
    right_space = 0,                   # plot_area_space="abs"のときの右のスペース指定
) -> list[Any]:
    """シート・ブック内、または選択中のグラフを一括変更します。

    Parameters
    ----------
    target : {"sheet", "book", "selected", "dir"}
        "sheet": ws のグラフ、"book": wb の全ワークシートのグラフ。
        "selected": アクティブなExcelで選択中のグラフのみ。
        Ctrlを押しながらグラフをクリックして複数(1つ以上)選択してください。
        このモードでは ws・wb を指定できません。
        省略した対象は、アクティブなブック・シートから取得します。
    ws : xlwings.Sheet, optional
        対象シート。target="book" では指定できません。
    wb : xlwings.Book, optional
        対象ブック。ws と同時には指定できません。
    directory : str or pathlib.Path, optional
        target="dir" の対象フォルダ。xlsx/xlsm/xlsb/xls を処理し、上書き保存します。
        Excelの一時ファイル（~$で始まる名前）は除外します。
        対象一覧を表示し、yes入力時のみ実行します。それ以外は空リストを返します。
    recursive : bool, default False
        Trueでサブフォルダも対象にします。target="dir" 専用です。
    書式設定引数
        ModifyChart と共通の設定引数（preset、width_cm、title_font_size など）。
        NS の省略時は各グラフの現在の系列数を使用します。
        系列のデータ範囲指定に使う ws は、各グラフの所属シートになります。

    Returns
    -------
    list
        通常は処理した xlwings.Chart のリスト。グラフがなければ空リスト。
        target="dir" はグラフを処理・保存したファイルの絶対パス（Path）のリスト。

    Notes
    -----
    ModifyChart と同じく、主に Windows Excel の散布図が対象です。
    独立したグラフシートは対象外です。dirモードのみ上書き保存します。
    dirモードは専用のExcelを起動し、マクロ・イベント・リンク更新を無効にします。
    エラー時は現在のファイルを保存せず閉じ、停止します。先に保存したファイルは残ります。
    例外が発生した時点で停止します。先に適用した変更は元に戻りません。
    ModifyChart 内部で捕捉されるエラーの扱いは、元の関数と同じです。
    """
    # 公開引数を既存の一括処理へ渡す設定辞書にまとめる。
    kwargs = locals().copy()
    for key in ("target", "ws", "wb", "directory", "recursive"):
        kwargs.pop(key)
    if NS is None:
        kwargs.pop("NS")

    if target not in ("sheet", "book", "selected", "dir"):
        raise ValueError('target は "sheet", "book", "selected", "dir" のいずれかを指定してください。')
    if target == "dir":
        if ws is not None or wb is not None:
            raise ValueError('target="dir" では ws・wb を指定できません。')
        if directory is None:
            raise ValueError('target="dir" では directory を指定してください。')
    elif directory is not None or recursive:
        raise ValueError('directory・recursive は target="dir" 専用です。')
    if target == "selected" and (ws is not None or wb is not None):
        raise ValueError('target="selected" では ws・wb を指定できません。')
    if ws is not None and wb is not None:
        raise ValueError("ws と wb は同時に指定できません。")
    if target == "book" and ws is not None:
        raise ValueError('target="book" では wb を指定してください。')
    if "chart" in kwargs:
        raise TypeError("chart は指定できません。対象は target で指定してください。")
    # 引数名の誤りは、Excelを操作する前に検出する。
    signature(ModifyChart).bind(chart=None, ws=None, **kwargs)

    if target == "dir":
        folder = Path(directory).expanduser().resolve()
        if not folder.is_dir():
            raise NotADirectoryError(f"フォルダが見つかりません: {folder}")
        candidates = folder.rglob("*") if recursive else folder.iterdir()
        files = sorted(
            (p for p in candidates if p.is_file()
             and p.suffix.lower() in {".xlsx", ".xlsm", ".xlsb", ".xls"}
             and not p.name.startswith("~$")),
            key=lambda p: str(p).lower(),
        )
        if not files:
            print("対象のExcelファイルがありません。")
            return []
        print(f"対象ファイル（{len(files)}件・上書き保存）:")
        for index, path in enumerate(files, start=1):
            print(f"  {index}. {path}")
        ans = input("\n上記ファイルの全てにグラフの書式設定を適用しますか？[yes/N] : ")
        if ans.strip().lower() != "yes":
            print("処理を中止しました。")
            return []
        import xlwings as xw
        saved = []
        with xw.App(visible=False, add_book=False) as app:
            app.display_alerts = False
            app.enable_events = False
            app.api.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
            for path in files:
                book = None
                try:
                    book = app.books.open(str(path), update_links=False, read_only=False,
                                          ignore_read_only_recommended=True)
                    if book.api.ReadOnly:
                        raise PermissionError(f"読み取り専用で開かれました: {path}")
                    charts = BatchModifyCharts(target="book", wb=book, **kwargs)
                    if charts:
                        book.save()
                        saved.append(path)
                except Exception as exc:
                    raise RuntimeError(f"Excelファイルの処理に失敗しました: {path}") from exc
                finally:
                    if book is not None:
                        book.close()
        return saved

    if target == "selected":
        import xlwings as xw
        app = xw.apps.active
        if app is None or app.books.active is None:
            raise RuntimeError("アクティブなExcelブックがありません。")
        sheet = app.books.active.sheets.active
        try:
            shapes = app.api.Selection.ShapeRange
        except Exception as exc:
            raise ValueError(
                "Ctrlを押しながらグラフを選択してください。(複数可)"
            ) from exc
        # ModifyChart による選択解除・名前変更の前に対象を固定する。
        names = [
            shapes.Item(i).Name
            for i in range(1, shapes.Count + 1)
            if shapes.Item(i).HasChart == -1
        ]
        jobs = [(sheet, sheet.charts[name]) for name in names]
        
    else:  # "sheet", "book" の場合
        if ws is not None:
            sheets = [ws]
        else:
            if wb is None:
                import xlwings as xw
                wb = xw.books.active
                if wb is None:
                    raise RuntimeError("アクティブなExcelブックがありません。")
            if target == "sheet":
                sheets = [wb.sheets.active]
            elif target == "book":
                sheets = list(wb.sheets)
            else:
                # 冒頭でも検証するが、未知の値をbook扱いにしない。
                raise ValueError(f"不明なtargetです: {target!r}")
        # 名前変更などでコレクションが変化しても、列挙対象を固定する。
        jobs = [(sheet, chart) for sheet in sheets for chart in list(sheet.charts)]
    modified = []
    for sheet, chart in jobs:
        options = deepcopy(kwargs)
        if "NS" not in options:
            options["NS"] = chart.api[1].SeriesCollection().Count
        # ModifyChart が series_list を変更するため、グラフごとにコピーする。
        ModifyChart(chart=chart, ws=sheet, **options)
        modified.append(chart)
        
    return modified
