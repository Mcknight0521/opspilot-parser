import io
import os
import re
from datetime import datetime
from typing import Any

import fitz  # PyMuPDF
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI(title="OpsPilot PDF Parser", version="1.0.0")

allowed_origins = [
    "https://mcknight0521.github.io",
    "http://localhost:8000",
    "http://127.0.0.1:8000",
]
extra_origin = os.getenv("OPSPILOT_FRONTEND_ORIGIN", "").strip()
if extra_origin and extra_origin not in allowed_origins:
    allowed_origins.append(extra_origin)

app.add_middleware(
    CORSMiddleware,
    allow_origins=allowed_origins,
    allow_credentials=False,
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["*"],
)

DATE_RE = re.compile(r"\b(\d{2})/(\d{2})/(\d{4})\b")
SKU_RE = re.compile(r"(?m)^([0-9]{6})$")
REF_RE = re.compile(r"F\d{6,}")
REASON_RE = re.compile(r"\b([A-Z]-[A-Za-z][A-Za-z-]*)\b")
NUM_RE = re.compile(r"^-?[\d,]+(?:\.\d+)?$")
TIME_RE = re.compile(r"^\d{2}:\d{2}:\d{2}$")


def clean(s: Any) -> str:
    return re.sub(r"\s+", " ", str(s or "").replace("\u00a0", " ")).strip()


def to_num(s: str | None) -> float | None:
    if s is None:
        return None
    try:
        return float(str(s).replace(",", ""))
    except Exception:
        return None


def iso_dmy(s: str | None) -> str | None:
    if not s:
        return None
    try:
        return datetime.strptime(s, "%d/%m/%Y").strftime("%Y-%m-%d")
    except Exception:
        return None


def extract_period(text: str) -> dict[str, str | None]:
    # Prefer the report's explicit adjustment-date range.
    m = re.search(
        r"調整日\s*(\d{2}/\d{2}/\d{4})\s*迄\s*(\d{2}/\d{2}/\d{4})",
        text,
        flags=re.S,
    )
    if not m:
        # Fallback to the first plausible date pair after a report heading.
        ds = DATE_RE.findall(text)
        if len(ds) >= 2:
            a = "/".join(ds[0])
            b = "/".join(ds[1])
            return {"start": iso_dmy(a), "end": iso_dmy(b)}
        return {"start": None, "end": None}
    return {"start": iso_dmy(m.group(1)), "end": iso_dmy(m.group(2))}


def cluster_blocks(blocks: list[tuple], rotation: int) -> list[list[tuple]]:
    """Group transaction blocks by their record axis.

    These ERP PDFs are often rotated 90 degrees. In that case each transaction is a
    vertical strip sharing x0; otherwise records are horizontal strips sharing y0.
    """
    axis = 0 if rotation in (90, 270) else 1
    candidates = []
    for b in blocks:
        x0, y0, x1, y1, text, *_ = b
        if not clean(text):
            continue
        coord = float(x0 if axis == 0 else y0)
        candidates.append((coord, b))
    candidates.sort(key=lambda z: z[0])

    groups: list[list[tuple]] = []
    centers: list[float] = []
    tol = 1.25
    for coord, b in candidates:
        idx = None
        for i, c in enumerate(centers):
            if abs(coord - c) <= tol:
                idx = i
                break
        if idx is None:
            groups.append([b])
            centers.append(coord)
        else:
            groups[idx].append(b)
            n = len(groups[idx])
            centers[idx] = ((centers[idx] * (n - 1)) + coord) / n
    return groups


def parse_inventory_adjustment(pdf: fitz.Document, filename: str) -> dict[str, Any]:
    full_text = "\n".join(page.get_text("text", sort=False) for page in pdf)
    if not re.search(r"庫存\s*調整\s*單", full_text):
        raise ValueError("not_inventory_adjustment")

    period = extract_period(full_text)
    rows: list[dict[str, Any]] = []
    reasons: set[str] = set()
    raw_candidates = 0

    for page_no, page in enumerate(pdf, start=1):
        blocks = page.get_text("blocks", sort=False)
        page_text = "\n".join(str(b[4]) for b in blocks)
        reason_match = REASON_RE.search(page_text)
        reason = reason_match.group(1) if reason_match else ""
        if reason:
            reasons.add(reason)

        for group in cluster_blocks(blocks, page.rotation):
            # Preserve physical order within the record strip.
            group_sorted = sorted(group, key=lambda b: (b[1], b[0]))
            full = "\n".join(str(b[4]).strip() for b in group_sorted if str(b[4]).strip())
            sku_match = SKU_RE.search(full)
            if not sku_match:
                continue
            raw_candidates += 1
            sku = sku_match.group(1)
            ref_match = REF_RE.search(full)
            dates = [m.group(0) for m in DATE_RE.finditer(full)]
            if not ref_match or len(dates) < 1:
                continue

            # Product name: the Chinese line immediately after HH:MM:SS in the same block.
            item = ""
            for b in group:
                lines = [x.strip() for x in str(b[4]).splitlines() if x.strip()]
                for i, token in enumerate(lines[:-1]):
                    if TIME_RE.fullmatch(token):
                        item = clean(lines[i + 1])
                        break
                if item:
                    break

            # Find the block carrying SKU. Its numeric prefix is consistently:
            # adjustment amount, stock before, adjustment qty, detail code, SKU, date.
            sku_block_lines: list[str] = []
            for b in group:
                lines = [x.strip() for x in str(b[4]).splitlines() if x.strip()]
                if sku in lines:
                    sku_block_lines = lines
                    break
            try:
                sku_idx = sku_block_lines.index(sku)
            except ValueError:
                continue

            detail = ""
            if sku_idx > 0 and re.fullmatch(r"\d{3}", sku_block_lines[sku_idx - 1]):
                detail = sku_block_lines[sku_idx - 1]
            prefix = sku_block_lines[: max(0, sku_idx - 1)]
            prefix_nums = [to_num(x) for x in prefix if NUM_RE.fullmatch(x)]
            prefix_nums = [x for x in prefix_nums if x is not None]
            if len(prefix_nums) < 3:
                continue
            adjustment_amount = prefix_nums[-3]
            stock_before = prefix_nums[-2]
            adjustment_qty = prefix_nums[-1]

            # Find block containing department 20 + Y/N. Its trailing numeric pattern is:
            # stock after, clearance flag, diff %, unit cost, department 20.
            unit_cost = None
            stock_after = None
            diff_pct = None
            clearance_flag = ""
            english_name = ""
            for b in group:
                lines = [x.strip() for x in str(b[4]).splitlines() if x.strip()]
                if "20" not in lines or not any(x in ("Y", "N") for x in lines):
                    continue
                vals = [to_num(x) for x in lines if NUM_RE.fullmatch(x)]
                vals = [x for x in vals if x is not None]
                if vals and abs(vals[-1] - 20) < 1e-9 and len(vals) >= 4:
                    unit_cost = vals[-2]
                    diff_pct = vals[-3]
                    stock_after = vals[-4]
                    clearance_flag = next((x for x in lines if x in ("Y", "N")), "")
                    # First nonnumeric/nonflag line is commonly English product name.
                    english_name = next(
                        (
                            clean(x)
                            for x in lines
                            if not NUM_RE.fullmatch(x)
                            and x not in ("Y", "N")
                            and x != "20"
                        ),
                        "",
                    )
                    break

            if adjustment_qty is None or unit_cost is None:
                continue
            if adjustment_amount is None:
                adjustment_amount = round(unit_cost * adjustment_qty, 2)

            adjust_date = iso_dmy(dates[0])
            post_date = iso_dmy(dates[-1]) if dates else adjust_date
            is_loss_reason = reason.startswith("E-") or reason.startswith("B-")
            waste = round(-adjustment_amount, 2) if is_loss_reason and adjustment_amount < 0 else None

            rows.append(
                {
                    "date": adjust_date,
                    "sku": sku,
                    "item": item or english_name or sku,
                    "englishName": english_name or None,
                    "qty": None,
                    "sales": None,
                    "waste": waste,
                    "clearance": None,
                    "periodStart": period["start"],
                    "periodEnd": period["end"],
                    "adjustmentReason": reason or "未辨識",
                    "adjustmentQty": adjustment_qty,
                    "adjustmentAmount": round(adjustment_amount, 2),
                    "unitCost": unit_cost,
                    "stockBefore": stock_before,
                    "stockAfter": stock_after,
                    "detailCode": detail,
                    "clearanceFlag": clearance_flag,
                    "diffPct": diff_pct,
                    "ref": ref_match.group(0),
                    "postDate": post_date,
                    "sourceType": "api-pdf-inventory-adjustment",
                    "sourceFile": filename,
                    "sourcePage": page_no,
                }
            )

    if not rows:
        raise ValueError("inventory_adjustment_unparsed")

    coverage = len(rows) / raw_candidates if raw_candidates else 1.0
    return {
        "ok": True,
        "documentType": "inventory_adjustment",
        "rows": rows,
        "meta": {
            "parser": "backend-pymupdf-inventory-v1",
            "reportType": "庫存調整單",
            "sourceType": "api-pdf",
            "sourceFile": filename,
            "parsedTransactions": len(rows),
            "rawTransactionCount": raw_candidates,
            "coverage": round(coverage, 4),
            "adjustmentReasons": sorted(reasons),
            "periodStart": period["start"],
            "periodEnd": period["end"],
        },
    }


@app.get("/health")
def health() -> dict[str, Any]:
    return {"ok": True, "service": "opspilot-parser", "version": "1.0.0"}


@app.post("/parse-pdf")
async def parse_pdf(file: UploadFile = File(...)) -> dict[str, Any]:
    name = file.filename or "upload.pdf"
    if not name.lower().endswith(".pdf"):
        raise HTTPException(status_code=415, detail="只接受 PDF 檔案")
    data = await file.read()
    if not data:
        raise HTTPException(status_code=400, detail="空白 PDF")
    if len(data) > 25 * 1024 * 1024:
        raise HTTPException(status_code=413, detail="PDF 超過 25MB")

    try:
        pdf = fitz.open(stream=data, filetype="pdf")
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"PDF 無法開啟：{e}")

    try:
        result = parse_inventory_adjustment(pdf, name)
        result["meta"]["pages"] = len(pdf)
        return result
    except ValueError as e:
        code = str(e)
        if code == "not_inventory_adjustment":
            raise HTTPException(status_code=422, detail="目前後端解析器尚未支援這種 PDF；前端可改用既有解析器。")
        raise HTTPException(status_code=422, detail="已辨識為庫存調整單，但交易明細重建失敗。")
    finally:
        pdf.close()
