#!/usr/bin/env python3
"""
Downloads the live workbook from OneDrive by following the share link redirects.
Works with standard "Anyone with the link" OneDrive shares.
"""
import base64
import json
import os
import sys
import tempfile
import urllib.parse
import requests
from datetime import datetime, timezone

import openpyxl

ONEDRIVE_SHARE_URL = os.environ["ONEDRIVE_SHARE_URL"]
SHEET_NAME = os.environ.get("SHEET_NAME")

DAY_COL_START = 2
DAY_COL_END = 32


def _looks_like_xlsx(content: bytes) -> bool:
    """xlsx files are zip archives — real ones start with the 'PK' signature.
    If OneDrive hands us an HTML viewer page instead, this catches it before
    openpyxl tries (and fails) to parse it."""
    return content[:2] == b"PK"


def _onedrive_api_url(share_url: str) -> str:
    b64 = base64.urlsafe_b64encode(share_url.strip().encode("utf-8")).decode("utf-8")
    b64 = b64.rstrip("=")
    return f"https://api.onedrive.com/v1.0/shares/u!{b64}/root/content"


def _looks_like_xlsx(content: bytes) -> bool:
    """xlsx files are zip archives — real ones start with the 'PK' signature.
    If OneDrive hands us an HTML viewer page instead, this catches it before
    openpyxl tries (and fails) to parse it."""
    return content[:2] == b"PK"


def _onedrive_api_url(share_url: str) -> str:
    b64 = base64.urlsafe_b64encode(share_url.strip().encode("utf-8")).decode("utf-8")
    b64 = b64.rstrip("=")
    return f"https://api.onedrive.com/v1.0/shares/u!{b64}/root/content"


def _with_query_param(url: str, key: str, value: str) -> str:
    """Return `url` with `key` set to `value` in the query string, replacing
    any existing value for that key."""
    parts = urllib.parse.urlsplit(url)
    query = urllib.parse.parse_qsl(parts.query, keep_blank_values=True)
    query = [(k, v) for k, v in query if k != key]
    query.append((key, value))
    new_query = urllib.parse.urlencode(query)
    return urllib.parse.urlunsplit((parts.scheme, parts.netloc, parts.path, new_query, parts.fragment))


def download_workbook(share_url: str) -> str:
    """Download the workbook from a public 'Anyone with the link' OneDrive
    share, trying several methods and validating the result actually looks
    like an .xlsx file before handing it to openpyxl. Raises a clear error
    if every method fails, instead of silently corrupting data.json."""
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"})
    tmp_path = os.path.join(tempfile.gettempdir(), "meal_chart_live.xlsx")
    errors = []

    def try_get(url, label):
        try:
            resp = session.get(url, timeout=60, allow_redirects=True)
            if resp.ok and _looks_like_xlsx(resp.content):
                with open(tmp_path, "wb") as f:
                    f.write(resp.content)
                return tmp_path
            errors.append(f"{label}: HTTP {resp.status_code}, "
                           f"content looked like xlsx = {_looks_like_xlsx(resp.content)}")
        except Exception as e:
            errors.append(f"{label}: {e}")
        return None

    # Method 1: resolve the share link, then — if it landed on OneDrive's
    # "Doc.aspx" web viewer page — swap action=default for action=download,
    # which makes OneDrive stream the raw file instead of the viewer.
    try:
        resp = session.get(share_url, allow_redirects=True, timeout=30)
        final_url = resp.url
        if _looks_like_xlsx(resp.content):
            with open(tmp_path, "wb") as f:
                f.write(resp.content)
            return tmp_path

        if "Doc.aspx" in final_url or "action=" in final_url:
            result = try_get(_with_query_param(final_url, "action", "download"),
                              "Doc.aspx action=download method")
            if result:
                return result

        # Older-style onedrive.live.com links respond to a bare download=1 flag.
        result = try_get(_with_query_param(final_url, "download", "1"),
                          "download=1 method")
        if result:
            return result
    except Exception as e:
        errors.append(f"redirect-resolve step: {e}")

    # Last resort: the legacy anonymous-share API. Only works for some
    # personal Microsoft accounts — many now return 401 for it.
    result = try_get(_onedrive_api_url(share_url), "api.onedrive.com method")
    if result:
        return result

    raise RuntimeError(
        "Could not download a valid .xlsx from ONEDRIVE_SHARE_URL. "
        "All download methods failed:\n  - " + "\n  - ".join(errors) +
        "\nCheck that the link is still shared as 'Anyone with the link can view' "
        "and hasn't expired or been replaced."
    )


def to_float(v, default=0.0):
    try:
        if v is None or str(v).strip() in ("", "\u200b"):
            return default
        return float(v)
    except (ValueError, TypeError):
        return default


def norm(s):
    return " ".join(str(s).split()).strip().lower() if s is not None else ""


def find_label_value(ws, label, max_cols_right=6, occurrence=0):
    matches = []
    for row in ws.iter_rows():
        for cell in row:
            if norm(cell.value) == norm(label):
                matches.append(cell)
    if not matches or occurrence >= len(matches):
        return None
    label_cell = matches[occurrence]
    for offset in range(1, max_cols_right + 1):
        v = ws.cell(row=label_cell.row, column=label_cell.column + offset).value
        if v is not None:
            return v
    return None


def parse_people(ws, bazar_start_row):
    people = []
    r = 2
    while r < bazar_start_row:
        name_cell = ws.cell(row=r, column=1).value
        name = str(name_cell).strip() if name_cell is not None else ""
        if not name or name.upper() == "N/A":
            r += 4
            continue

        def day_counts(row):
            return {
                str(c - DAY_COL_START + 1): to_float(ws.cell(row=row, column=c).value)
                for c in range(DAY_COL_START, DAY_COL_END + 1)
            }

        breakfast = day_counts(r)
        lunch = day_counts(r + 1)
        dinner = day_counts(r + 2)

        days = {}
        for d in breakfast.keys():
            b, l, din = breakfast[d], lunch.get(d, 0.0), dinner.get(d, 0.0)
            days[d] = {"breakfast": b, "lunch": l, "dinner": din, "total": b * 0.5 + l + din}

        people.append({"name": name, "row": r, "days": days})
        r += 4
    return people


def find_header_columns(ws, header_row=1):
    wanted = {
        "personal total": "total_meals",
        "personal cost(tk)": "personal_cost",
        "deposit(tk)": "deposit",
        "balance(tk)": "balance",
        "khalar bill": "khalar",
        "gas bill": "gas",
        "electricity bill": "electricity",
    }
    cols = {}
    for cell in ws[header_row]:
        key = norm(cell.value)
        if key in wanted:
            cols[wanted[key]] = cell.column
    return cols


def find_bazar_section(ws):
    title_cell = None
    for row in ws.iter_rows():
        for cell in row:
            if norm(cell.value) == "bazar list":
                title_cell = cell
                break
        if title_cell:
            break
    if not title_cell:
        return [], None

    header_row = title_cell.row + 1
    col_map = {}
    for cell in ws[header_row]:
        key = norm(cell.value)
        if key in ("date", "cost", "person", "main items"):
            col_map[key] = cell.column

    bazar = []
    r = header_row + 1
    blank_streak = 0
    while blank_streak < 3 and r < ws.max_row:
        date_val = ws.cell(row=r, column=col_map.get("date", 4)).value
        if date_val is None or str(date_val).strip() == "":
            blank_streak += 1
            r += 1
            continue
        blank_streak = 0
        date_str = date_val.strftime("%Y-%m-%d") if hasattr(date_val, "strftime") else str(date_val).strip()
        if "." in date_str and len(date_str.split(".")) == 3:
            dd, mm, yyyy = date_str.split(".")
            date_str = f"{yyyy}-{mm}-{dd}"
        bazar.append({
            "date": date_str,
            "cost": to_float(ws.cell(row=r, column=col_map.get("cost", 9)).value),
            "person": str(ws.cell(row=r, column=col_map.get("person", 14)).value or "").strip(),
            "items": str(ws.cell(row=r, column=col_map.get("main items", 20)).value or "").strip(),
        })
        r += 1
    return bazar, title_cell.row


def main():
    xlsx_path = download_workbook(ONEDRIVE_SHARE_URL)
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb[SHEET_NAME] if SHEET_NAME else wb.worksheets[0]

    bazar, bazar_title_row = find_bazar_section(ws)
    bazar.sort(key=lambda b: b["date"])

    header_cols = find_header_columns(ws)
    people_raw = parse_people(ws, bazar_title_row or ws.max_row)

    people = []
    grand_total_meal = 0.0
    for p in people_raw:
        r = p["row"]
        total_meals = ws.cell(row=r, column=header_cols["total_meals"]).value if "total_meals" in header_cols else sum(d["total"] for d in p["days"].values())
        total_meals = to_float(total_meals, default=sum(d["total"] for d in p["days"].values()))
        grand_total_meal += total_meals

        def col_val(key, fallback=0.0):
            if key in header_cols:
                return to_float(ws.cell(row=r, column=header_cols[key]).value, fallback)
            return fallback

        people.append({
            "name": p["name"],
            "days": p["days"],
            "total_meals": total_meals,
            "personal_cost": round(col_val("personal_cost"), 2),
            "deposit": col_val("deposit"),
            "balance": round(col_val("balance"), 2),
            "khalar": col_val("khalar"),
            "gas": col_val("gas"),
            "electricity": col_val("electricity"),
        })

    meal_rate = to_float(find_label_value(ws, "Meal Rate"))
    total_bazar_cost = to_float(find_label_value(ws, "Total (Cost) Bazar")) or sum(b["cost"] for b in bazar)
    total_deposit = to_float(find_label_value(ws, "Total Deposit")) or sum(p["deposit"] for p in people)
    total_due = to_float(find_label_value(ws, "Total due"))
    current_balance = to_float(find_label_value(ws, "Current Balance"))
    grand_total_meal_label = to_float(find_label_value(ws, "Grand Total (Meal)"))
    if grand_total_meal_label:
        grand_total_meal = grand_total_meal_label

    data = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "meal_rate": round(meal_rate, 4),
        "grand_total_meal": grand_total_meal,
        "total_bazar_cost": round(total_bazar_cost, 2),
        "total_deposit": round(total_deposit, 2),
        "total_due": round(total_due, 2) if total_due else round(total_deposit - sum(p["personal_cost"] for p in people), 2),
        "current_balance": round(current_balance, 2),
        "people": sorted(people, key=lambda p: p["name"]),
        "bazar": bazar,
    }

    with open("data.json", "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

    print(f"Wrote data.json — {len(people)} people, {len(bazar)} bazar entries, "
          f"meal_rate={data['meal_rate']}, grand_total_meal={data['grand_total_meal']}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)
