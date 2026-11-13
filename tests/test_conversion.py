import io

from openpyxl import load_workbook

from app.services import conversion_service as cs

from .conftest import make_xlsx


def test_csv_to_xlsx_keeps_leading_zero_and_formula_text():
    raw = "id;nama;nilai\n001;=SUM(A1);12\n002;Budi;7.5\n".encode("utf-8-sig")
    data, notes = cs.csv_to_xlsx(raw, "auto", "auto", "auto")
    ws = load_workbook(io.BytesIO(data)).active
    assert ws["A2"].value == "001"
    assert ws["B2"].value == "=SUM(A1)" and ws["B2"].data_type == "s"
    assert ws["C2"].value == 12
    assert any("leading zero" in n for n in notes)


def test_csv_text_mode_keeps_everything_text():
    data, _ = cs.csv_to_xlsx(b"a,b\n1,2\n", "auto", "auto", "text")
    ws = load_workbook(io.BytesIO(data)).active
    assert ws["A2"].value == "1"


def test_xlsx_to_csv_selected_sheet():
    raw = make_xlsx([["a", "b"], [1, 2.5], [None, "x"]], sheet="Lembar")
    data, notes = cs.xlsx_to_csv(raw, "Lembar", "semicolon", "utf-8")
    assert data.decode().splitlines() == ["a;b", "1;2.5", ";x"]
    assert "tidak dipertahankan" in notes[0]
