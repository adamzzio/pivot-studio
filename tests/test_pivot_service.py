import math

import pandas as pd
import pytest

from app.services import excel_service as xs
from app.services import pivot_service as ps
from app.utils import AppError

from .conftest import SALES


def frame():
    df = pd.DataFrame(SALES[1:], columns=SALES[0]).astype(object)
    df = df.where(df.notna(), float("nan"))
    inferred = {c: xs.infer_type(df[c]) for c in df.columns}
    return df, inferred


def cfg(**kw):
    base = {"source": {"file_id": "x" * 32, "sheet": "Data"}, "rows": [], "columns": [], "values": [], "filters": []}
    base.update(kw)
    return ps.parse_config(base)


def val(result, kind, labels, idx=0):
    for b in result["body"]:
        if b["kind"] == kind and b["labels"] == labels:
            return b["values"][idx]
    raise AssertionError("baris tidak ditemukan")


def run(**kw):
    df, inf = frame()
    return ps.compute_pivot(df, inf, cfg(**kw))


def test_type_inference_keeps_leading_zero_as_text():
    _, inf = frame()
    assert inf["Sales"] == "number" and inf["Code"] == "text" and inf["Region"] == "text"


def test_sum_with_subtotals_and_grand_total():
    r = run(rows=["Region", "Product"], values=[{"field": "Sales", "agg": "SUM"}])
    assert val(r, "detail", ["A", "X"]) == 40
    assert val(r, "subtotal", ["A"]) == 60
    assert val(r, "subtotal", ["B"]) == 100
    assert val(r, "grand", []) == 160


def test_count_all_rows_vs_non_null():
    r = run(rows=["Region"], values=[{"field": "Sales", "agg": "COUNT"}, {"field": "Sales", "agg": "COUNT_NON_NULL"}])
    assert val(r, "detail", ["B"], 0) == 3
    assert val(r, "detail", ["B"], 1) == 2


def test_variance_and_std_are_population():
    r = run(rows=["Region"], values=[{"field": "Sales", "agg": "VAR"}, {"field": "Sales", "agg": "STD"}])
    # Region A: 10, 20, 30 -> var populasi = 66.666..., bukan 100 (sample)
    assert math.isclose(val(r, "detail", ["A"], 0), 200 / 3)
    assert math.isclose(val(r, "detail", ["A"], 1), math.sqrt(200 / 3))


def test_median_distinct_and_mode_tie_picks_smallest():
    r = run(rows=["Region"], values=[
        {"field": "Sales", "agg": "MEDIAN"},
        {"field": "Product", "agg": "COUNT_DISTINCT"},
        {"field": "Qty", "agg": "MODE"},
    ])
    assert val(r, "detail", ["A"], 0) == 20
    assert val(r, "detail", ["A"], 1) == 2
    assert val(r, "detail", ["A"], 2) == 2          # Qty A: 1,2,2
    assert val(r, "detail", ["B"], 2) == 4          # Qty B: 4,4,5
    assert val(r, "grand", [], 2) == 2              # seri 2 dan 4, dipilih yang terkecil
    assert any("MODE" in n for n in r["meta"]["notes"])


def test_mode_no_repeat_returns_none():
    r = run(rows=["Product"], values=[{"field": "Sales", "agg": "MODE"}])
    assert val(r, "detail", ["X"]) is None


def test_columns_cross_tab_with_grand_column():
    r = run(rows=["Region"], columns=["Product"], values=[{"field": "Sales", "agg": "SUM"}])
    assert r["header"]["col_paths"] == [["X"], ["Y"]]
    assert r["header"]["grand_col"] is True
    row_a = next(b for b in r["body"] if b["labels"] == ["A"])
    assert row_a["values"] == [40, 20, 60]


def test_numeric_agg_on_text_field_is_rejected():
    with pytest.raises(AppError):
        run(rows=["Region"], values=[{"field": "Product", "agg": "SUM"}])


def test_filters_values_and_range():
    r = run(rows=["Region"], values=[{"field": "Sales", "agg": "SUM"}],
            filters=[{"field": "Region", "kind": "values", "values": ["A"]}])
    assert r["meta"]["records"] == 3 and val(r, "grand", []) == 60
    r = run(rows=["Region"], values=[{"field": "Qty", "agg": "SUM"}],
            filters=[{"field": "Qty", "kind": "range", "min": 4, "max": None}])
    assert r["meta"]["records"] == 3


def test_sort_by_value_desc():
    r = run(rows=["Region"], values=[{"field": "Sales", "agg": "SUM"}],
            options={"row_sort": {"by": "value", "dir": "desc", "metric": 0}})
    labels = [b["labels"] for b in r["body"] if b["kind"] == "detail"]
    assert labels == [["B"], ["A"]]


def test_field_cannot_be_in_rows_and_columns():
    with pytest.raises(AppError):
        run(rows=["Region"], columns=["Region"], values=[{"field": "Sales", "agg": "SUM"}])


def test_values_required():
    with pytest.raises(AppError):
        run(rows=["Region"])
