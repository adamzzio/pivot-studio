"""Integrasi Groq API. Dipanggil hanya dari backend; API key dari environment variable."""
from __future__ import annotations

import json
import re

from app.utils import AppError

LANGS = {"id": "Bahasa Indonesia", "en": "English"}
LENGTHS = {
    "ringkas": "ringkas: maksimal 3 butir per bagian dan ringkasan 2 sampai 3 kalimat",
    "detail": "detail: hingga 6 butir per bagian dan ringkasan 4 sampai 6 kalimat",
}
SECTIONS = ["findings", "extremes", "comparisons", "anomalies", "recommendations", "limitations"]


def _fmt(v) -> str:
    if v is None:
        return "-"
    if isinstance(v, float):
        return f"{v:.4f}".rstrip("0").rstrip(".")
    return str(v)


def build_payload(result: dict, cfg: dict, max_rows: int, max_cols: int) -> tuple[str, bool]:
    hdr, meta = result["header"], result["meta"]
    metrics = hdr["metrics"]
    names = []
    for cp in hdr["col_paths"]:
        for m in metrics:
            names.append(" / ".join(list(cp) + [m["label"]]))
    if hdr["grand_col"]:
        names += [f"Grand Total / {m['label']}" for m in metrics]
    truncated = False
    if len(names) > max_cols:
        names, truncated = names[:max_cols], True

    body = result["body"]
    if len(body) > max_rows:
        grand = [b for b in body if b["kind"] == "grand"]
        body = body[: max_rows - len(grand)] + grand
        truncated = True

    lines = [
        "Konfigurasi pivot:",
        f"- Baris: {', '.join(hdr['row_fields']) or '(tidak ada)'}",
        f"- Kolom: {', '.join(hdr['col_fields']) or '(tidak ada)'}",
        "- Nilai: " + "; ".join(f"{m['label']} = {m['agg']} dari {m['field']}" for m in metrics),
        f"- Record dianalisis: {meta['records']} dari {meta['total_records']}",
        "",
        "Data pivot (satu baris per entri; format: jenis | label baris | kolom=nilai):",
    ]
    for b in body:
        label = " > ".join(b["labels"]) if b["labels"] else "Total"
        vals = "; ".join(f"{n}={_fmt(v)}" for n, v in zip(names, b["values"]))
        lines.append(f"{b['kind']} | {label} | {vals}")
    if truncated:
        lines.append("")
        lines.append("CATATAN: data dipotong demi ukuran; sebagian baris atau kolom tidak disertakan.")
    return "\n".join(lines), truncated


def _system_prompt(language: str, length: str) -> str:
    return (
        "Kamu adalah analis data. Analisis HANYA data pivot yang diberikan pengguna.\n"
        "Aturan wajib:\n"
        "1. Jangan mengarang angka. Setiap angka yang kamu sebut harus tertulis persis di data.\n"
        "2. Jika data tidak cukup untuk suatu kesimpulan, nyatakan eksplisit di bagian limitations.\n"
        "3. Baris berjenis subtotal dan grand adalah total, bukan kategori terpisah.\n"
        "4. Balas HANYA dengan JSON valid tanpa teks lain, dengan kunci: "
        "title (string), summary (string), findings (array string), extremes (array string, nilai tertinggi dan terendah), "
        "comparisons (array string, perbandingan antar kategori), anomalies (array string), "
        "recommendations (array string), limitations (array string).\n"
        f"5. Bahasa output: {LANGS[language]}.\n"
        f"6. Panjang: {LENGTHS[length]}."
    )


def _parse(text: str) -> dict | None:
    t = (text or "").strip()
    t = re.sub(r"^```(?:json)?\s*|\s*```$", "", t, flags=re.I)
    try:
        data = json.loads(t)
    except ValueError:
        m = re.search(r"\{.*\}", t, re.S)
        if not m:
            return None
        try:
            data = json.loads(m.group(0))
        except ValueError:
            return None
    if not isinstance(data, dict):
        return None
    out = {"title": str(data.get("title") or "Insight Pivot Table")[:200], "summary": str(data.get("summary") or "")}
    for k in SECTIONS:
        v = data.get(k)
        out[k] = [str(x) for x in v][:20] if isinstance(v, list) else ([str(v)] if v else [])
    return out


def generate_insight(payload_text: str, *, api_key: str, model: str, language: str, length: str, timeout: float) -> dict:
    if not api_key:
        raise AppError("GROQ_API_KEY belum diatur. Isi pada file .env lalu jalankan ulang aplikasi.", 503)
    try:
        import groq
    except ImportError as exc:
        raise AppError("Paket 'groq' belum terpasang. Jalankan: pip install -r requirements.txt", 503) from exc

    client = groq.Groq(api_key=api_key, timeout=timeout, max_retries=1)
    messages = [
        {"role": "system", "content": _system_prompt(language, length)},
        {"role": "user", "content": payload_text},
    ]
    kwargs = dict(model=model, messages=messages, temperature=0.2, max_tokens=4096,
                  extra_body={"reasoning_effort": "low"})
    try:
        try:
            resp = client.chat.completions.create(response_format={"type": "json_object"}, **kwargs)
        except groq.BadRequestError:
            resp = client.chat.completions.create(**kwargs)
    except groq.APITimeoutError as exc:
        raise AppError("Permintaan ke Groq melewati batas waktu. Coba lagi.", 504) from exc
    except groq.RateLimitError as exc:
        raise AppError("Batas penggunaan Groq tercapai. Tunggu sebentar lalu coba lagi.", 429) from exc
    except groq.AuthenticationError as exc:
        raise AppError("API key Groq tidak valid. Periksa GROQ_API_KEY di file .env.", 502) from exc
    except groq.APIConnectionError as exc:
        raise AppError("Tidak dapat terhubung ke Groq. Periksa koneksi internet.", 502) from exc
    except groq.APIStatusError as exc:
        raise AppError(f"Groq mengembalikan error (kode {exc.status_code}). Coba lagi.", 502) from exc

    content = (resp.choices[0].message.content or "") if resp.choices else ""
    parsed = _parse(content)
    if parsed:
        return {"insight": parsed}
    if not content.strip():
        raise AppError("Groq tidak mengembalikan hasil. Coba lagi.", 502)
    return {"raw_text": content[:8000]}
