import csv
import io
import os
import re
from datetime import datetime
from pathlib import Path
from typing import Any

import fitz
import pandas as pd
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook

APP_VERSION = "3.0.0"
MAX_BYTES = 30 * 1024 * 1024
SUPPORTED = ["pdf", "xlsx", "xls", "xlsm", "ods", "csv", "tsv", "txt"]

app = FastAPI(title="OpsPilot Universal Parser", version=APP_VERSION)
origins = [
    "https://mcknight0521.github.io",
    "http://localhost:8000",
    "http://127.0.0.1:8000",
]
extra = os.getenv("OPSPILOT_FRONTEND_ORIGIN", "").strip()
if extra and extra not in origins:
    origins.append(extra)
app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=False,
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["*"],
)

DATE_DMY_RE = re.compile(r"\b(\d{2})/(\d{2})/(\d{4})\b")
DATE_YMD_RE = re.compile(r"\b(\d{4})[-/](\d{1,2})[-/](\d{1,2})\b")
SKU_RE = re.compile(r"(?m)^([0-9]{6})$")
REF_RE = re.compile(r"F\d{6,}")
REASON_RE = re.compile(r"\b([A-Z]-[A-Za-z][A-Za-z-]*)\b")
NUM_RE = re.compile(r"^-?[\d,]+(?:\.\d+)?$")
TIME_RE = re.compile(r"^\d{2}:\d{2}:\d{2}$")

ALIASES = {
    "date": ["日期", "交易日期", "銷售日期", "調整日", "date", "sales date"],
    "sku": ["單品編號", "商品編號", "品號", "料號", "sku", "item code", "product code"],
    "item": ["單品名稱", "商品名稱", "品名", "名稱", "item name", "product name"],
    "qty": ["銷售量", "數量", "銷量", "qty", "quantity", "sales qty"],
    "sales": ["銷售額(含稅)", "含稅營業額", "營業額", "銷售額", "sales", "revenue", "amount"],
    "salesExTax": ["銷售額(未稅)", "未稅營業額", "未稅銷售額", "sales ex tax", "net sales"],
    "tax": ["稅額", "tax", "vat"],
    "waste": ["報廢", "報廢金額", "損耗", "損耗金額", "waste", "scrap"],
    "clearance": ["降價出清金額", "出清金額", "出清", "clearance", "markdown amount"],
    "clearanceQty": ["降價出清數量", "出清數量", "clearance qty", "markdown qty"],
}


def clean(v: Any) -> str:
    return re.sub(r"\s+", " ", str(v if v is not None else "").replace("\u00a0", " ")).strip()


def norm_header(v: Any) -> str:
    return re.sub(r"[\s　]+", "", clean(v)).lower()


def num(v: Any) -> float:
    if v is None or clean(v) == "":
        return 0.0
    s = clean(v).replace(",", "").replace("$", "").replace("％", "%")
    if s.endswith("%"):
        s = s[:-1]
    try:
        return float(s)
    except Exception:
        return 0.0


def to_iso(v: Any):
    if v is None or clean(v) == "":
        return None
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    s = clean(v)
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except Exception:
            pass
    try:
        dt = pd.to_datetime(v, errors="coerce")
        if not pd.isna(dt):
            return dt.strftime("%Y-%m-%d")
    except Exception:
        pass
    return None


def close(a, b, tol=.02):
    try:
        return abs(float(a) - float(b)) <= max(tol, abs(float(b)) * 1e-7)
    except Exception:
        return False


def vcheck(name, source, calc, tol=.02):
    if source is None:
        return {"metric": name, "status": "unverified", "sourceTotal": None, "recalculatedTotal": round(calc, 2), "difference": None}
    ok = close(calc, source, tol)
    return {
        "metric": name,
        "status": "passed" if ok else "failed",
        "sourceTotal": round(float(source), 2),
        "recalculatedTotal": round(float(calc), 2),
        "difference": round(float(calc) - float(source), 2),
    }


def summarize_checks(checks, **extra):
    failed = [c for c in checks if c.get("status") == "failed"]
    partial = [c for c in checks if c.get("status") in ("partial", "unverified")]
    status = "failed" if failed else ("partial" if partial else "passed")
    out = {
        "status": status,
        "checks": checks,
        "passed": sum(c.get("status") == "passed" for c in checks),
        "failed": len(failed),
        "unverified": sum(c.get("status") == "unverified" for c in checks),
        "warning": sum(c.get("status") == "partial" for c in checks),
    }
    out.update(extra)
    return out


def extract_period_text(text: str):
    m = re.search(r"調整日\s*(\d{2}/\d{2}/\d{4})\s*迄\s*(\d{2}/\d{2}/\d{4})", text, re.S)
    if m:
        return {"start": to_iso(m.group(1)), "end": to_iso(m.group(2))}
    dates = [m.group(0) for m in DATE_DMY_RE.finditer(text)]
    if len(dates) >= 2:
        # Printed date often appears first; use earliest and latest unique date when possible.
        iso = [to_iso(x) for x in dates]
        iso = sorted({x for x in iso if x})
        if len(iso) >= 2:
            return {"start": iso[0], "end": iso[-1]}
    return {"start": None, "end": None}


def cluster_blocks(blocks, rotation):
    axis = 0 if rotation in (90, 270) else 1
    items = [(float(b[axis]), b) for b in blocks if clean(b[4])]
    items.sort(key=lambda z: z[0])
    groups, centers = [], []
    for coord, b in items:
        idx = next((i for i, c in enumerate(centers) if abs(coord - c) <= 1.25), None)
        if idx is None:
            groups.append([b])
            centers.append(coord)
        else:
            groups[idx].append(b)
            n = len(groups[idx])
            centers[idx] = (centers[idx] * (n - 1) + coord) / n
    return groups


def parse_inventory_pdf(data: bytes, name: str):
    pdf = fitz.open(stream=data, filetype="pdf")
    try:
        full = "\n".join(p.get_text("text", sort=False) for p in pdf)
        if not re.search(r"庫存\s*調整\s*單", full):
            raise ValueError("unsupported")
        period = extract_period_text(full)
        rows, reasons, raw = [], set(), 0
        for page_no, page in enumerate(pdf, start=1):
            blocks = page.get_text("blocks", sort=False)
            page_text = "\n".join(str(b[4]) for b in blocks)
            rm = REASON_RE.search(page_text)
            reason = rm.group(1) if rm else ""
            if reason:
                reasons.add(reason)
            for group in cluster_blocks(blocks, page.rotation):
                fullg = "\n".join(str(b[4]).strip() for b in sorted(group, key=lambda b: (b[1], b[0])) if clean(b[4]))
                sm = SKU_RE.search(fullg)
                if not sm:
                    continue
                raw += 1
                sku = sm.group(1)
                ref = REF_RE.search(fullg)
                dates = [m.group(0) for m in DATE_DMY_RE.finditer(fullg)]
                if not ref or not dates:
                    continue
                item = ""
                for b in group:
                    lines = [x.strip() for x in str(b[4]).splitlines() if x.strip()]
                    for i, t in enumerate(lines[:-1]):
                        if TIME_RE.fullmatch(t):
                            item = clean(lines[i + 1])
                            break
                    if item:
                        break
                sku_lines = []
                for b in group:
                    lines = [x.strip() for x in str(b[4]).splitlines() if x.strip()]
                    if sku in lines:
                        sku_lines = lines
                        break
                try:
                    si = sku_lines.index(sku)
                except Exception:
                    continue
                detail = sku_lines[si - 1] if si > 0 and re.fullmatch(r"\d{3}", sku_lines[si - 1]) else ""
                vals = [num(x) for x in sku_lines[: max(0, si - 1)] if NUM_RE.fullmatch(x)]
                if len(vals) < 3:
                    continue
                amount, stock_before, qty = vals[-3:]
                unit = stock_after = diff = None
                flag = en = ""
                for b in group:
                    lines = [x.strip() for x in str(b[4]).splitlines() if x.strip()]
                    if "20" not in lines or not any(x in ("Y", "N") for x in lines):
                        continue
                    vs = [num(x) for x in lines if NUM_RE.fullmatch(x)]
                    if len(vs) >= 4 and abs(vs[-1] - 20) < 1e-9:
                        unit, diff, stock_after = vs[-2], vs[-3], vs[-4]
                        flag = next((x for x in lines if x in ("Y", "N")), "")
                        en = next((clean(x) for x in lines if not NUM_RE.fullmatch(x) and x not in ("Y", "N") and x != "20"), "")
                        break
                if unit is None:
                    continue
                waste = round(-amount, 2) if (reason.startswith("E-") or reason.startswith("B-")) and amount < 0 else None
                rows.append({
                    "date": to_iso(dates[0]), "sku": sku, "item": item or en or sku, "englishName": en or None,
                    "qty": None, "sales": None, "waste": waste, "clearance": None,
                    "periodStart": period["start"], "periodEnd": period["end"],
                    "adjustmentReason": reason or "未辨識", "adjustmentQty": qty,
                    "adjustmentAmount": round(amount, 2), "unitCost": unit,
                    "stockBefore": stock_before, "stockAfter": stock_after, "detailCode": detail,
                    "clearanceFlag": flag, "diffPct": diff, "ref": ref.group(0), "postDate": to_iso(dates[-1]),
                    "sourceType": "api-pdf-inventory-adjustment", "sourceFile": name, "sourcePage": page_no,
                })
        if not rows:
            raise ValueError("unparsed")
        math_ok = sum(1 for r in rows if close(r["adjustmentAmount"], r["adjustmentQty"] * r["unitCost"], .05))
        checks = [{
            "metric": "adjustment_amount_math",
            "status": "passed" if math_ok == len(rows) else "partial",
            "matchedRows": math_ok,
            "totalRows": len(rows),
        }]
        validation = summarize_checks(checks, rawTransactionCount=raw, parsedTransactions=len(rows), coverage=round(len(rows) / raw, 4) if raw else 1)
        return {
            "ok": True, "documentType": "inventory_adjustment", "confidence": 1.0 if validation["status"] == "passed" else .95,
            "validation": validation, "rows": rows,
            "meta": {"parser": "backend-pymupdf-inventory-v3", "reportType": "庫存調整單", "sourceType": "api-pdf", "sourceFile": name,
                     "pages": len(pdf), "parsedTransactions": len(rows), "rawTransactionCount": raw,
                     "coverage": round(len(rows) / raw, 4) if raw else 1, "adjustmentReasons": sorted(reasons),
                     "periodStart": period["start"], "periodEnd": period["end"]},
        }
    finally:
        pdf.close()


def parse_sales_pdf(data: bytes, name: str):
    pdf = fitz.open(stream=data, filetype="pdf")
    try:
        full = "\n".join(p.get_text("text", sort=False) for p in pdf)
        if "每日銷售報表 - 依單品" not in full:
            raise ValueError("unsupported")
        period = extract_period_text(full)
        # The last page grand total is a stable audit anchor for this report family.
        mtotal = re.search(r"總計\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)\s+(-?[\d,]+)\s+([\d,.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)", full, re.S)
        source_total = None
        if mtotal:
            vals = [num(x) for x in mtotal.groups()]
            # Extracted order: ex-tax, tax, incl-tax, stock, qty, margin, benefit, immediate qty, immediate amt, clearance qty, clearance amt
            source_total = {"salesExTax": vals[0], "tax": vals[1], "sales": vals[2], "qty": vals[4], "clearanceQty": vals[9], "clearance": vals[10]}

        rows_raw = []
        for page_no, page in enumerate(pdf, start=1):
            lines = [x.strip() for x in page.get_text("text", sort=False).splitlines() if x.strip()]
            i = 0
            while i < len(lines):
                if not re.fullmatch(r"\d{6}", lines[i]):
                    i += 1
                    continue
                sku = lines[i]
                # Do not parse if this SKU is immediately part of a header-like block; require N/P marker before next SKU.
                j = i + 1
                while j < len(lines) and not re.fullmatch(r"\d{6}", lines[j]) and lines[j] not in ("小分類總計", "總計"):
                    j += 1
                block = lines[i + 1:j]
                np_idx = next((k for k, x in enumerate(block) if x in ("N", "P")), None)
                if np_idx is None or np_idx < 9 or len(block) < np_idx + 5:
                    i = max(j, i + 1)
                    continue
                pre = block[:np_idx]
                post = block[np_idx + 1:]
                # Expect final numeric scaffold before N/P: qty, ex-tax, tax, sales, GM, benefit, stock, dept, detail
                nums_pre = [x for x in pre if NUM_RE.fullmatch(x)]
                nums_post = [x for x in post if NUM_RE.fullmatch(x)]
                if len(nums_pre) < 9 or len(nums_post) < 4:
                    i = max(j, i + 1)
                    continue
                n = [num(x) for x in nums_pre[-9:]]
                p = [num(x) for x in nums_post[:4]]
                qty, ex_tax, tax, sales, gm, benefit, stock, dept, detail = n
                immediate_qty, immediate_amt, clearance_qty, clearance_amt = p
                # Name text is everything before the numeric scaffold, with Chinese/English separated heuristically.
                first_num_pos = next((k for k, x in enumerate(pre) if NUM_RE.fullmatch(x)), len(pre))
                names = pre[:first_num_pos]
                item = names[0] if names else sku
                if len(names) > 1 and any('\u4e00' <= ch <= '\u9fff' for ch in names[1]):
                    # wrapped Chinese product name
                    item = "".join(x for x in names if any('\u4e00' <= ch <= '\u9fff' for ch in x)) or item
                english = " ".join(x for x in names if not any('\u4e00' <= ch <= '\u9fff' for ch in x)) or None
                rows_raw.append({
                    "date": None, "sku": sku, "item": clean(item), "englishName": clean(english) if english else None, "np": block[np_idx],
                    "qty": qty, "sales": sales, "salesExTax": ex_tax, "tax": tax, "waste": None,
                    "clearance": clearance_amt, "clearanceQty": clearance_qty, "immediateClearanceQty": immediate_qty,
                    "immediateClearanceAmount": immediate_amt, "grossMarginPct": gm, "commercialBenefitPct": benefit,
                    "stock": stock, "periodStart": period["start"], "periodEnd": period["end"],
                    "sourceType": "api-pdf-daily-sales", "sourceFile": name, "sourcePage": page_no,
                })
                i = max(j, i + 1)
        if not rows_raw:
            raise ValueError("unparsed")
        result = finalize_daily_sales(rows_raw, name, "api-pdf", None, source_total, excluded_summary_rows=None)
        result["meta"]["parser"] = "backend-pymupdf-daily-sales-v3"
        result["meta"]["pages"] = len(pdf)
        return result
    finally:
        pdf.close()


def split_item(v):
    parts = [x.strip() for x in str(v or "").splitlines() if x.strip()]
    return (parts[0] if parts else "", " ".join(parts[1:]) if len(parts) > 1 else None)


def find_header(table):
    required = {norm_header(x) for x in ["單品編號", "單品名稱", "銷售量", "銷售額(含稅)"]}
    for i, row in enumerate(table[:150]):
        vals = {norm_header(x) for x in row if clean(x)}
        if required.issubset(vals):
            return i
    return None


def header_map(header):
    return {norm_header(h): i for i, h in enumerate(header) if clean(h)}


def hidx(idx, name):
    return idx.get(norm_header(name))


def finalize_daily_sales(detail, name, source_type, sheet, source_totals=None, excluded_summary_rows=0):
    merged = {}
    for r in detail:
        k = r["sku"]
        if k not in merged:
            merged[k] = {**r, "np": None, "qty": 0.0, "sales": 0.0, "salesExTax": 0.0, "tax": 0.0, "clearance": 0.0, "clearanceQty": 0.0}
        m = merged[k]
        for f in ("qty", "sales", "salesExTax", "tax", "clearance", "clearanceQty"):
            m[f] = (m.get(f) or 0) + (r.get(f) or 0)
    rows = list(merged.values())
    calc = {
        "qty": sum(r.get("qty") or 0 for r in rows),
        "sales": sum(r.get("sales") or 0 for r in rows),
        "salesExTax": sum(r.get("salesExTax") or 0 for r in rows),
        "tax": sum(r.get("tax") or 0 for r in rows),
        "clearance": sum(r.get("clearance") or 0 for r in rows),
        "clearanceQty": sum(r.get("clearanceQty") or 0 for r in rows),
    }
    checks = []
    for metric in ("sales", "qty", "clearance"):
        checks.append(vcheck(metric, source_totals.get(metric) if source_totals else None, calc[metric], 1.05 if metric == "sales" else .05))
    if source_totals and source_totals.get("salesExTax") is not None and source_totals.get("tax") is not None:
        checks.append(vcheck("sales_ex_tax", source_totals.get("salesExTax"), calc["salesExTax"], 1.05))
        checks.append(vcheck("tax", source_totals.get("tax"), calc["tax"], 1.05))
    tax_bad = sum(1 for r in rows if not close((r.get("salesExTax") or 0) + (r.get("tax") or 0), r.get("sales") or 0, 1.05))
    checks.append({"metric": "row_tax_math", "status": "passed" if tax_bad == 0 else "failed", "matchedRows": len(rows) - tax_bad, "totalRows": len(rows)})
    dup_removed = len(detail) - len(rows)
    validation = summarize_checks(
        checks,
        excludedSummaryRows=excluded_summary_rows,
        rawDetailRows=len(detail),
        mergedSkuRows=len(rows),
        duplicateNpRowsMerged=dup_removed,
    )
    if validation["status"] == "failed":
        raise HTTPException(status_code=422, detail={"message": "報表驗算失敗，為避免錯誤數字，本次不匯入。", "validation": validation})
    return {
        "ok": True,
        "documentType": "daily_sales_by_item",
        "confidence": 1.0 if validation["status"] == "passed" else .90,
        "validation": validation,
        "totals": {k: round(v, 2) for k, v in calc.items()},
        "rows": rows,
        "meta": {
            "parser": "backend-daily-sales-v3", "reportType": "每日銷售報表 - 依單品", "sourceType": source_type,
            "sourceFile": name, "sheet": sheet, "periodStart": rows[0].get("periodStart") if rows else None,
            "periodEnd": rows[0].get("periodEnd") if rows else None, "rawDetailRows": len(detail), "parsedProducts": len(rows),
            "excludedSummaryRows": excluded_summary_rows,
        },
    }


def parse_daily_sales_table(table, name, source_type, sheet=None):
    hi = find_header(table)
    if hi is None:
        raise ValueError("unsupported")
    header = [clean(x) for x in table[hi]]
    idx = header_map(header)
    needed = ["單品編號", "單品名稱", "銷售量", "銷售額(含稅)", "降價出清金額"]
    if any(hidx(idx, x) is None for x in needed):
        raise ValueError("unsupported")
    start = end = None
    for row in table[:hi]:
        vals = [clean(x) for x in row]
        # Robustly collect dates in the report header; use earliest/latest.
        dates = [to_iso(x) for x in vals]
        dates = [x for x in dates if x]
        if dates:
            start = min([start] + dates) if start else min(dates)
            end = max([end] + dates) if end else max(dates)
    detail, summary_rows, source_totals = [], 0, None
    for row in table[hi + 1:]:
        vals = [clean(x) for x in row]
        if any(x in ("小分類總計", "總計") for x in vals):
            summary_rows += 1
            if "總計" in vals and "小分類總計" not in vals:
                source_totals = {
                    "qty": num(row[hidx(idx, "銷售量")]), "sales": num(row[hidx(idx, "銷售額(含稅)")]),
                    "salesExTax": num(row[hidx(idx, "銷售額(未稅)")]) if hidx(idx, "銷售額(未稅)") is not None else None,
                    "tax": num(row[hidx(idx, "稅額")]) if hidx(idx, "稅額") is not None else None,
                    "clearance": num(row[hidx(idx, "降價出清金額")]),
                    "clearanceQty": num(row[hidx(idx, "降價出清數量")]) if hidx(idx, "降價出清數量") is not None else None,
                }
            continue
        si = hidx(idx, "單品編號")
        sku = clean(row[si]) if si is not None and si < len(row) else ""
        if not re.fullmatch(r"\d{6}", sku):
            continue
        item, en = split_item(row[hidx(idx, "單品名稱")])
        np_i = hidx(idx, "N/P")
        detail.append({
            "date": None, "sku": sku, "item": item or sku, "englishName": en,
            "np": clean(row[np_i]) if np_i is not None else None,
            "qty": num(row[hidx(idx, "銷售量")]), "sales": num(row[hidx(idx, "銷售額(含稅)")]),
            "salesExTax": num(row[hidx(idx, "銷售額(未稅)")]) if hidx(idx, "銷售額(未稅)") is not None else 0,
            "tax": num(row[hidx(idx, "稅額")]) if hidx(idx, "稅額") is not None else 0,
            "waste": None, "clearance": num(row[hidx(idx, "降價出清金額")]),
            "clearanceQty": num(row[hidx(idx, "降價出清數量")]) if hidx(idx, "降價出清數量") is not None else 0,
            "periodStart": start, "periodEnd": end, "sourceType": source_type, "sourceFile": name,
        })
    if not detail:
        raise ValueError("unparsed")
    return finalize_daily_sales(detail, name, source_type, sheet, source_totals, summary_rows)


def dataframe_to_table(df: pd.DataFrame):
    return [[None if pd.isna(x) else x for x in row] for row in df.itertuples(index=False, name=None)]


def load_spreadsheet_tables(data: bytes, ext: str):
    if ext in (".xlsx", ".xlsm"):
        wb = load_workbook(io.BytesIO(data), data_only=True, read_only=False)
        return [(ws.title, [list(r) for r in ws.iter_rows(values_only=True)]) for ws in wb.worksheets]
    if ext == ".xls":
        sheets = pd.read_excel(io.BytesIO(data), sheet_name=None, header=None, engine="xlrd")
        return [(name, dataframe_to_table(df)) for name, df in sheets.items()]
    if ext == ".ods":
        sheets = pd.read_excel(io.BytesIO(data), sheet_name=None, header=None, engine="odf")
        return [(name, dataframe_to_table(df)) for name, df in sheets.items()]
    raise ValueError("unsupported")


def decode_text(data: bytes):
    for enc in ("utf-8-sig", "cp950", "big5", "utf-8", "utf-16", "latin1"):
        try:
            return data.decode(enc), enc
        except Exception:
            pass
    raise ValueError("encoding")


def text_to_table(data: bytes, ext: str):
    text, enc = decode_text(data)
    if ext == ".tsv":
        delim = "\t"
    elif ext == ".csv":
        sample = text[:20000]
        try:
            delim = csv.Sniffer().sniff(sample, delimiters=",;\t|").delimiter
        except Exception:
            delim = ","
    else:
        sample = text[:20000]
        try:
            delim = csv.Sniffer().sniff(sample, delimiters=",;\t|").delimiter
        except Exception:
            # TXT with fixed-ish spacing: convert 2+ spaces to tabs as a conservative fallback.
            if any("\t" in x for x in text.splitlines()[:20]):
                delim = "\t"
            else:
                text = "\n".join(re.sub(r" {2,}", "\t", ln.rstrip()) for ln in text.splitlines())
                delim = "\t"
    return list(csv.reader(io.StringIO(text), delimiter=delim)), enc, delim


def match_alias(header_value: Any):
    h = norm_header(header_value)
    if not h:
        return None
    for field, aliases in ALIASES.items():
        for a in aliases:
            if h == norm_header(a):
                return field
    return None


def find_generic_header(table):
    best = None
    for i, row in enumerate(table[:150]):
        mapping = {}
        for j, v in enumerate(row):
            f = match_alias(v)
            if f and f not in mapping:
                mapping[f] = j
        score = len(mapping)
        if score >= 2 and ("sales" in mapping or "waste" in mapping or "clearance" in mapping or "qty" in mapping):
            if best is None or score > best[0]:
                best = (score, i, mapping)
    return best


def parse_generic_table(table, name, source_type, sheet=None):
    found = find_generic_header(table)
    if not found:
        raise ValueError("unsupported")
    score, hi, mapping = found
    if "sku" not in mapping and "item" not in mapping:
        raise ValueError("unsupported")
    rows = []
    excluded = 0
    for row in table[hi + 1:]:
        vals = [clean(x) for x in row]
        if any(re.search(r"(^|\s)(小計|總計|subtotal|grand total)(\s|$)", x, re.I) for x in vals if x):
            excluded += 1
            continue
        sku = clean(row[mapping["sku"]]) if "sku" in mapping and mapping["sku"] < len(row) else ""
        item = clean(row[mapping["item"]]) if "item" in mapping and mapping["item"] < len(row) else ""
        if not sku and not item:
            continue
        sales = num(row[mapping["sales"]]) if "sales" in mapping else None
        qty = num(row[mapping["qty"]]) if "qty" in mapping else None
        waste = num(row[mapping["waste"]]) if "waste" in mapping else None
        clearance = num(row[mapping["clearance"]]) if "clearance" in mapping else None
        date = to_iso(row[mapping["date"]]) if "date" in mapping else None
        rows.append({
            "date": date, "sku": sku or item, "item": item or sku, "englishName": None,
            "qty": qty, "sales": sales, "salesExTax": num(row[mapping["salesExTax"]]) if "salesExTax" in mapping else None,
            "tax": num(row[mapping["tax"]]) if "tax" in mapping else None,
            "waste": waste, "clearance": clearance,
            "clearanceQty": num(row[mapping["clearanceQty"]]) if "clearanceQty" in mapping else None,
            "sourceType": source_type, "sourceFile": name,
        })
    if not rows:
        raise ValueError("unparsed")
    dates = sorted({r["date"] for r in rows if r.get("date")})
    for r in rows:
        r["periodStart"] = dates[0] if dates else None
        r["periodEnd"] = dates[-1] if dates else None
    checks = []
    if "salesExTax" in mapping and "tax" in mapping and "sales" in mapping:
        checked = [r for r in rows if r.get("sales") is not None]
        bad = sum(1 for r in checked if not close((r.get("salesExTax") or 0) + (r.get("tax") or 0), r.get("sales") or 0, 1.05))
        checks.append({"metric": "row_tax_math", "status": "passed" if bad == 0 else "partial", "matchedRows": len(checked)-bad, "totalRows": len(checked)})
    checks.append({"metric": "source_grand_total", "status": "unverified", "reason": "通用格式未找到可安全辨識的原始總計；已保留為部分驗證。"})
    confidence = min(.88, .50 + score * .06)
    validation = summarize_checks(checks, mappedFields=sorted(mapping.keys()), excludedSummaryRows=excluded, parsedRows=len(rows))
    totals = {
        "sales": round(sum(r.get("sales") or 0 for r in rows), 2) if "sales" in mapping else None,
        "qty": round(sum(r.get("qty") or 0 for r in rows), 2) if "qty" in mapping else None,
        "waste": round(sum(r.get("waste") or 0 for r in rows), 2) if "waste" in mapping else None,
        "clearance": round(sum(r.get("clearance") or 0 for r in rows), 2) if "clearance" in mapping else None,
    }
    return {
        "ok": True, "documentType": "generic_operational_table", "confidence": confidence,
        "validation": validation, "totals": totals, "rows": rows,
        "meta": {"parser": "backend-generic-table-v3", "reportType": "通用營運表格", "sourceType": source_type,
                 "sourceFile": name, "sheet": sheet, "periodStart": dates[0] if dates else None, "periodEnd": dates[-1] if dates else None,
                 "mappedFields": sorted(mapping.keys()), "requiresReview": True},
    }


def parse_tabular(data: bytes, name: str, ext: str):
    if ext in (".xlsx", ".xlsm", ".xls", ".ods"):
        tables = load_spreadsheet_tables(data, ext)
        source_type = f"api-{ext[1:]}"
    else:
        table, enc, delim = text_to_table(data, ext)
        tables = [(None, table)]
        source_type = f"api-{ext[1:] or 'txt'}"
    # First pass: known report fingerprints. Second pass: generic mapping.
    errors = []
    for sheet, table in tables:
        try:
            return parse_daily_sales_table(table, name, source_type, sheet)
        except HTTPException:
            raise
        except Exception as e:
            errors.append(f"{sheet or 'text'}:known:{e}")
    candidates = []
    for sheet, table in tables:
        try:
            candidates.append(parse_generic_table(table, name, source_type, sheet))
        except Exception as e:
            errors.append(f"{sheet or 'text'}:generic:{e}")
    if candidates:
        return sorted(candidates, key=lambda x: x.get("confidence", 0), reverse=True)[0]
    raise ValueError("unsupported")


def parse_any(data: bytes, name: str):
    ext = Path(name).suffix.lower()
    if ext == ".pdf":
        # Try known PDF report families in deterministic order.
        for parser in (parse_inventory_pdf, parse_sales_pdf):
            try:
                return parser(data, name)
            except HTTPException:
                raise
            except ValueError as e:
                if str(e) != "unsupported":
                    raise
        raise HTTPException(status_code=422, detail="PDF 已讀取，但目前無法安全辨識此報表；本次資料未匯入。")
    if ext in (".xlsx", ".xls", ".xlsm", ".ods", ".csv", ".tsv", ".txt"):
        return parse_tabular(data, name, ext)
    raise HTTPException(status_code=415, detail=f"目前接受：{', '.join(SUPPORTED).upper()}")


@app.get("/")
def root():
    return {"ok": True, "service": "OpsPilot Universal Parser", "version": APP_VERSION, "docs": "/docs", "health": "/health"}


@app.get("/health")
def health():
    return {"ok": True, "service": "opspilot-parser", "version": APP_VERSION, "formats": SUPPORTED, "endpoint": "/parse-file"}


@app.post("/parse-file")
async def parse_file(file: UploadFile = File(...)):
    name = file.filename or "upload"
    data = await file.read()
    if not data:
        raise HTTPException(400, "空白檔案")
    if len(data) > MAX_BYTES:
        raise HTTPException(413, "檔案超過 30MB")
    try:
        return parse_any(data, name)
    except HTTPException:
        raise
    except ValueError as e:
        if str(e) == "unsupported":
            raise HTTPException(422, "目前無法安全辨識此報表；可改用欄位確認流程建立新模板，本次不會把未確認數字寫入正式分析。")
        raise HTTPException(422, f"檔案已讀取，但驗證或解析失敗：{e}")
    except Exception as e:
        raise HTTPException(400, f"檔案無法開啟：{e}")


# Backward compatibility for the current website while it migrates to /parse-file.
@app.post("/parse-pdf")
async def parse_pdf(file: UploadFile = File(...)):
    name = file.filename or "upload.pdf"
    data = await file.read()
    if not name.lower().endswith(".pdf"):
        raise HTTPException(415, "只接受 PDF 檔案")
    try:
        return parse_any(data, name)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(422, f"PDF 解析失敗：{e}")
