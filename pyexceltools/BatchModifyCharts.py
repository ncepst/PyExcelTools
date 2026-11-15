"""複数の埋め込みグラフに ModifyChart の設定を適用します。"""
from copy import deepcopy
from inspect import signature
from typing import Any, Literal

from .ModifyChart import ModifyChart


def BatchModifyCharts(
    target: Literal["sheet", "book", "selected"] = "sheet",
    *,
    ws: Any = None,
    wb: Any = None,
    **kwargs: Any,
) -> list[Any]:
    """シート・ブック内、または選択中のグラフを一括変更します。

    Parameters
    ----------
    target : {"sheet", "book", "selected"}
        "sheet": ws のグラフ、"book": wb の全ワークシートのグラフ。
        "selected": アクティブなExcelで選択中のグラフのみ。
        Ctrlを押しながらグラフの外枠をクリックして複数選択してください。
        このモードでは ws・wb を指定できません。
        省略した対象は、アクティブなブック・シートから取得します。
    ws : xlwings.Sheet, optional
        対象シート。target="book" では指定できません。
    wb : xlwings.Book, optional
        対象ブック。ws と同時には指定できません。
    **kwargs
        ModifyChart の設定引数（preset、width_cm、title_font_size など）。
        NS の省略時は各グラフの現在の系列数を使用します。
        系列のデータ範囲指定に使う ws は、各グラフの所属シートになります。

    Returns
    -------
    list
        処理した xlwings.Chart のリスト。グラフがなければ空リスト。

    Notes
    -----
    ModifyChart と同じく、主に Windows Excel の散布図が対象です。
    独立したグラフシートは対象外です。保存は行いません。
    例外が発生した時点で停止します。先に適用した変更は元に戻りません。
    ModifyChart 内部で捕捉されるエラーの扱いは、元の関数と同じです。
    """
    if target not in ("sheet", "book", "selected"):
        raise ValueError('target は "sheet", "book", "selected" のいずれかを指定してください。')
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
                "グラフの外枠を選択してください。複数選択はCtrlを押しながらクリックします。"
            ) from exc
        # ModifyChart による選択解除・名前変更の前に対象を固定する。
        names = [
            shapes.Item(i).Name
            for i in range(1, shapes.Count + 1)
            if shapes.Item(i).HasChart == -1
        ]
        jobs = [(sheet, sheet.charts[name]) for name in names]
    else:
        if ws is not None:
            sheets = [ws]
        else:
            if wb is None:
                import xlwings as xw
                wb = xw.books.active
                if wb is None:
                    raise RuntimeError("アクティブなExcelブックがありません。")
            sheets = [wb.sheets.active] if target == "sheet" else list(wb.sheets)
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