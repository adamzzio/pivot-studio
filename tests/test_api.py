import io

from openpyxl import load_workbook

from .conftest import SALES, make_xlsx


def upload(client, rows=SALES, name="sales.xlsx"):
    r = client.post("/api/pivot/upload", data={"files": (io.BytesIO(make_xlsx(rows)), name)}, content_type="multipart/form-data")
    assert r.status_code == 200, r.get_json()
    return r.get_json()["files"][0]


def payload(f, **kw):
    base = {"source": {"file_id": f["file_id"], "sheet": "Data"}, "rows": ["Region"], "columns": [],
            "values": [{"field": "Sales", "agg": "SUM", "label": "Total"}], "filters": []}
    base.update(kw)
    return base


def test_pages_render(client):
    assert client.get("/").status_code == 302
    for p in ("/converter", "/pivot"):
        html = client.get(p).get_data(as_text=True)
        assert "Pivot Studio" in html


def test_upload_rejects_non_xlsx(client):
    r = client.post("/api/pivot/upload", data={"files": (io.BytesIO(b"abc"), "x.xlsx")}, content_type="multipart/form-data")
    assert r.status_code == 400 and "bukan XLSX" in r.get_json()["error"]


def test_preview_calculate_export_flow(client):
    f = upload(client)
    pv = client.post("/api/pivot/preview", json={"file_id": f["file_id"], "sheet": "Data"}).get_json()
    assert pv["total_rows"] == 6 and {c["name"]: c["type"] for c in pv["columns"]}["Code"] == "text"

    calc = client.post("/api/pivot/calculate", json=payload(f)).get_json()
    assert calc["ok"] and calc["result"]["body"][-1]["values"] == [160]

    ex = client.post("/api/pivot/export", json=payload(f))
    assert ex.status_code == 200
    wb = load_workbook(io.BytesIO(ex.data))
    assert wb.sheetnames == ["Pivot", "Raw Data"]
    assert wb["Raw Data"].max_row == 7
    texts = [str(c.value) for row in wb["Pivot"].iter_rows() for c in row if c.value is not None]
    assert "Grand Total" in texts and "Total" in texts


def test_field_values_search(client):
    f = upload(client)
    d = client.post("/api/pivot/field-values", json={"file_id": f["file_id"], "sheet": "Data", "field": "Product", "search": "y"}).get_json()
    assert [v["label"] for v in d["values"]] == ["Y"]


def test_session_isolation(app):
    c1, c2 = app.test_client(), app.test_client()
    f = upload(c1)
    r = c2.post("/api/pivot/preview", json={"file_id": f["file_id"], "sheet": "Data"})
    assert r.status_code == 404


def test_insights_requires_api_key(client):
    f = upload(client)
    r = client.post("/api/ai/insights", json=payload(f))
    assert r.status_code == 503 and "GROQ_API_KEY" in r.get_json()["error"]


def test_cross_origin_post_blocked(client):
    r = client.post("/api/pivot/clear", headers={"Origin": "http://evil.example"})
    assert r.status_code == 403


def test_converter_endpoints(client):
    r = client.post("/api/convert/csv-to-xlsx", data={"file": (io.BytesIO(b"a,b\n1,2\n"), "t.csv"), "type_mode": "auto"}, content_type="multipart/form-data")
    assert r.status_code == 200 and r.headers["Content-Disposition"].endswith('t.xlsx"') or "t.xlsx" in r.headers["Content-Disposition"]
    r = client.post("/api/convert/inspect", data={"kind": "xlsx-to-csv", "file": (io.BytesIO(make_xlsx([["a"]], "S1")), "t.xlsx")}, content_type="multipart/form-data")
    assert r.get_json()["info"]["sheets"][0]["name"] == "S1"
    r = client.post("/api/convert/xlsx-to-csv", data={"file": (io.BytesIO(b"bukan xlsx"), "t.xlsx"), "sheet": "S1"}, content_type="multipart/form-data")
    assert r.status_code == 400
