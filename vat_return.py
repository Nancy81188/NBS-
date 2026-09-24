"""Lebanese quarterly VAT return (Q1-Q4) built from sales, purchases, expenses and customs cases."""
from __future__ import annotations

import calendar
import json
from datetime import date
from decimal import Decimal, ROUND_HALF_UP

from database import display_date, iso_date, utcnow

ZERO = Decimal("0")
CENT = Decimal("0.01")
CATEGORIES = {"sales": "Sales", "purchases": "Purchases", "assets": "Fixed assets", "expenses": "Expenses", "customs": "Customs (import VAT)"}
ADJUSTMENT_TYPES = {"output": "Output VAT adjustment", "input": "Deductible VAT adjustment", "non_deductible": "Non-deductible VAT adjustment"}
LINES = (
    ("1", "sales", "Taxable sales - output VAT"),
    ("2", "adj_output", "Output VAT adjustments"),
    ("3", "total_output", "Total output VAT"),
    ("4", "purchases", "Deductible VAT - purchases"),
    ("5", "assets", "Deductible VAT - fixed assets"),
    ("6", "expenses", "Deductible VAT - expenses"),
    ("7", "customs", "Deductible VAT - customs / imports"),
    ("8", "adj_input", "Deductible VAT adjustments"),
    ("9", "total_input", "Total deductible VAT"),
    ("10", "non_deductible", "Non-deductible VAT (not claimed)"),
    ("11", "net", "Net VAT for the quarter (3 - 9)"),
)


def quarter_range(year, quarter):
    try: year = int(year); quarter = int(quarter)
    except (TypeError, ValueError) as exc: raise ValueError("Enter a valid year and quarter") from exc
    if year < 2000 or year > 2100 or quarter not in (1, 2, 3, 4): raise ValueError("Quarter must be Q1, Q2, Q3 or Q4")
    first = 3 * quarter - 2
    return date(year, first, 1).isoformat(), date(year, first + 2, calendar.monthrange(year, first + 2)[1]).isoformat()


def previous_quarter(year, quarter):
    return (int(year), int(quarter) - 1) if int(quarter) > 1 else (int(year) - 1, 4)


def _money(value):
    try: return Decimal(str(value or 0)).quantize(CENT, rounding=ROUND_HALF_UP)
    except Exception: return ZERO


def _lbp(value):
    return Decimal(str(value or 0)).quantize(Decimal("1"), rounding=ROUND_HALF_UP)


def _documents(db, start, end, currency, include_review):
    statuses = ("posted", "review") if include_review else ("posted",)
    documents = []; skipped = []; review_excluded = 0
    with db.connect() as connection:
        invoices = [dict(row) for row in connection.execute("""SELECT i.*,p.name party_name FROM invoices i
            LEFT JOIN parties p ON p.id=i.party_id WHERE i.status<>'cancelled'""")]
        expenses = [dict(row) for row in connection.execute("SELECT * FROM expenses")]
    for row in invoices:
        try: day = iso_date(row["invoice_date"])
        except ValueError: skipped.append(row["invoice_number"]); continue
        if not start <= day <= end or (currency and row["currency"] != currency): continue
        if row["status"] not in statuses:
            if row["status"] == "review": review_excluded += 1
            continue
        if row["kind"] == "sale": category = "sales"
        elif row.get("source_file") == "Customs Case": category = "customs"
        elif row.get("entry_type") == "assets": category = "assets"
        elif row.get("entry_type") == "expenses": category = "expenses"
        else: category = "purchases"
        documents.append({"source": "invoice", "id": row["id"], "date": day, "number": row["invoice_number"], "party": row.get("party_name") or "",
            "category": category, "currency": row["currency"], "base": _money(row.get("deductible_subtotal") or row.get("subtotal")),
            "exempt": _money(row.get("non_deductible_subtotal")), "vat": _money(row.get("vat")), "status": row["status"],
            "recoverable": category == "sales" or bool(int(row.get("vat_recoverable") if row.get("vat_recoverable") is not None else 1))})
    for row in expenses:
        try: day = iso_date(row["expense_date"])
        except ValueError: skipped.append(f"EXP-{row['id']}"); continue
        if not start <= day <= end or (currency and row["currency"] != currency): continue
        documents.append({"source": "expense", "id": row["id"], "date": day, "number": row.get("reference") or f"EXP-{row['id']}",
            "party": row.get("description") or "", "category": "expenses", "currency": row["currency"],
            "base": _money(row.get("with_vat_subtotal") or row.get("subtotal")), "exempt": _money(row.get("without_vat_subtotal")),
            "vat": _money(row.get("vat")), "status": "posted", "recoverable": bool(int(row.get("vat_recoverable") if row.get("vat_recoverable") is not None else 1))})
    documents.sort(key=lambda item: (item["date"], item["category"], str(item["number"])))
    return documents, skipped, review_excluded


def list_adjustments(db, year, quarter):
    with db.connect() as connection:
        return [dict(row) for row in connection.execute("SELECT * FROM vat_adjustments WHERE year=? AND quarter=? ORDER BY id", (int(year), int(quarter)))]


def saved_return(db, year, quarter):
    with db.connect() as connection:
        row = connection.execute("SELECT * FROM vat_returns WHERE year=? AND quarter=?", (int(year), int(quarter))).fetchone()
    return dict(row) if row else None


def list_saved_returns(db):
    with db.connect() as connection:
        return [dict(row) for row in connection.execute("""SELECT id,year,quarter,net_lbp,credit_brought_forward_lbp,payable_lbp,
            credit_carried_forward_lbp,saved_at,saved_by_name FROM vat_returns ORDER BY year DESC,quarter DESC""")]


def build_vat_return(db, year, quarter, currency=None, include_review=False, previous_year_db=None, credit_brought_forward=None):
    start, end = quarter_range(year, quarter)
    currency = str(currency).upper() if currency and str(currency).upper() not in ("ALL", "ALL CURRENCIES") else None
    documents, skipped, review_excluded = _documents(db, start, end, currency, include_review)
    rates = {}
    def to_lbp(amount, code, day):
        if code == "LBP": return _lbp(amount)
        key = (code, day)
        if key not in rates: rates[key] = db._converted_amount(Decimal("1"), code, "LBP", day)
        return _lbp(Decimal(amount) * rates[key])
    per_currency = {}
    def bucket(code):
        return per_currency.setdefault(code, {key: {"base": ZERO, "vat": ZERO, "vat_lbp": ZERO, "count": 0} for _, key, _ in LINES})
    for doc in documents:
        doc["vat_lbp"] = to_lbp(doc["vat"], doc["currency"], doc["date"])
        doc["lbp_rate"] = Decimal("1") if doc["currency"] == "LBP" else rates.get((doc["currency"], doc["date"]), Decimal("1"))
        key = doc["category"] if doc["recoverable"] else "non_deductible"
        line = bucket(doc["currency"])[key]
        line["base"] += doc["base"]; line["vat"] += doc["vat"]; line["vat_lbp"] += doc["vat_lbp"]; line["count"] += 1
    adjustments = list_adjustments(db, year, quarter)
    for adjustment in adjustments:
        if currency and adjustment["currency"] != currency: continue
        amount = _money(adjustment["amount"]); amount_lbp = to_lbp(amount, adjustment["currency"], end)
        adjustment["amount_lbp"] = amount_lbp
        key = {"output": "adj_output", "input": "adj_input", "non_deductible": "non_deductible"}[adjustment["adjustment_type"]]
        line = bucket(adjustment["currency"])[key]; line["vat"] += amount; line["vat_lbp"] += amount_lbp; line["count"] += 1
    for values in per_currency.values():
        for field in ("vat", "vat_lbp"):
            values["total_output"][field] = values["sales"][field] + values["adj_output"][field]
            values["total_input"][field] = sum((values[k][field] for k in ("purchases", "assets", "expenses", "customs", "adj_input")), ZERO)
            values["net"][field] = values["total_output"][field] - values["total_input"][field]
        values["total_output"]["base"] = values["sales"]["base"]
        values["total_input"]["base"] = sum((values[k]["base"] for k in ("purchases", "assets", "expenses", "customs")), ZERO)
    totals_lbp = {key: sum((values[key]["vat_lbp"] for values in per_currency.values()), ZERO) for _, key, _ in LINES}
    # Credit brought forward from the previous quarter's saved return (previous fiscal-year file for Q1).
    source = "none"
    if credit_brought_forward not in (None, ""):
        credit_bf = _lbp(credit_brought_forward); source = "manual"
    else:
        credit_bf = ZERO; previous_year, previous_q = previous_quarter(year, quarter)
        for candidate in ([db] if previous_year == int(year) else [previous_year_db, db]):
            if candidate is None: continue
            try: previous = saved_return(candidate, previous_year, previous_q)
            except Exception: previous = None
            if previous:
                credit_bf = _lbp(previous["credit_carried_forward_lbp"]); source = f"Q{previous_q} {previous_year} saved return"; break
    net_after_credit = totals_lbp["net"] - credit_bf
    payable = net_after_credit if net_after_credit > 0 else ZERO
    credit_cf = -net_after_credit if net_after_credit < 0 else ZERO
    saved = saved_return(db, year, quarter)
    changed = bool(saved) and (_lbp(saved["net_lbp"]) != _lbp(totals_lbp["net"]) or _lbp(saved["credit_brought_forward_lbp"]) != credit_bf)
    return {"year": int(year), "quarter": int(quarter), "date_from": start, "date_to": end, "currency_filter": currency or "All",
        "include_review": bool(include_review), "per_currency": per_currency, "totals_lbp": totals_lbp,
        "credit_brought_forward_lbp": credit_bf, "credit_source": source, "net_after_credit_lbp": net_after_credit,
        "payable_lbp": payable, "credit_carried_forward_lbp": credit_cf, "documents": documents, "adjustments": adjustments,
        "skipped": skipped, "review_excluded": review_excluded, "saved": saved, "changed_since_saved": changed,
        "status": "saved" if saved and not changed else "changed after saving" if changed else "not saved"}


def _ensure_not_saved(db, year, quarter):
    if saved_return(db, year, quarter):
        raise ValueError(f"The Q{int(quarter)} {int(year)} VAT return is saved. An administrator must reopen it before it can be changed")


def add_adjustment(db, item, user_id, user_name=""):
    year = int(item.get("year") or 0); quarter = int(item.get("quarter") or 0); quarter_range(year, quarter)
    kind = str(item.get("adjustment_type") or "").lower()
    if kind not in ADJUSTMENT_TYPES: raise ValueError("Choose Output, Deductible or Non-deductible VAT adjustment")
    currency = str(item.get("currency") or "LBP").upper()
    if currency not in ("USD", "EUR", "LBP", "AED"): raise ValueError("Invalid adjustment currency")
    try: amount = Decimal(str(item.get("amount") or "").replace(",", ""))
    except Exception as exc: raise ValueError("Adjustment amount must be a number (use a minus sign to reduce)") from exc
    if not amount: raise ValueError("Adjustment amount cannot be zero")
    reason = str(item.get("reason") or "").strip()
    if len(reason) < 3: raise ValueError("Enter the reason for the adjustment")
    _ensure_not_saved(db, year, quarter)
    with db.connect() as connection:
        adjustment_id = connection.execute("""INSERT INTO vat_adjustments(year,quarter,currency,adjustment_type,amount,reason,created_by,created_by_name,created_at)
            VALUES(?,?,?,?,?,?,?,?,?)""", (year, quarter, currency, kind, str(amount.quantize(CENT)), reason, user_id, user_name, utcnow())).lastrowid
        connection.execute("INSERT INTO audit_log(user_id,action,entity,entity_id,details,created_at) VALUES(?,?,?,?,?,?)",
            (user_id, "create", "vat_adjustment", adjustment_id, json.dumps({"year": year, "quarter": quarter, "type": kind, "amount": str(amount), "reason": reason}), utcnow()))
    return adjustment_id


def delete_adjustment(db, adjustment_id, user_id):
    with db.connect() as connection:
        row = connection.execute("SELECT * FROM vat_adjustments WHERE id=?", (int(adjustment_id),)).fetchone()
    if not row: raise KeyError("Adjustment not found")
    _ensure_not_saved(db, row["year"], row["quarter"])
    with db.connect() as connection:
        connection.execute("DELETE FROM vat_adjustments WHERE id=?", (int(adjustment_id),))
        connection.execute("INSERT INTO audit_log(user_id,action,entity,entity_id,details,created_at) VALUES(?,?,?,?,?,?)",
            (user_id, "delete", "vat_adjustment", int(adjustment_id), json.dumps(dict(row)), utcnow()))
    return {"deleted": int(adjustment_id)}


def save_return(db, year, quarter, user_id, previous_year_db=None, credit_brought_forward=None, user_name=""):
    result = build_vat_return(db, year, quarter, None, False, previous_year_db, credit_brought_forward)
    snapshot = json.dumps(json_ready({k: result[k] for k in ("per_currency", "totals_lbp", "credit_source", "review_excluded")}))
    with db.connect() as connection:
        connection.execute("""INSERT INTO vat_returns(year,quarter,net_lbp,credit_brought_forward_lbp,payable_lbp,credit_carried_forward_lbp,snapshot,saved_by,saved_by_name,saved_at)
            VALUES(?,?,?,?,?,?,?,?,?,?) ON CONFLICT(year,quarter) DO UPDATE SET net_lbp=excluded.net_lbp,
            credit_brought_forward_lbp=excluded.credit_brought_forward_lbp,payable_lbp=excluded.payable_lbp,
            credit_carried_forward_lbp=excluded.credit_carried_forward_lbp,snapshot=excluded.snapshot,saved_by=excluded.saved_by,
            saved_by_name=excluded.saved_by_name,saved_at=excluded.saved_at""",
            (int(year), int(quarter), str(result["totals_lbp"]["net"]), str(result["credit_brought_forward_lbp"]), str(result["payable_lbp"]),
             str(result["credit_carried_forward_lbp"]), snapshot, user_id, user_name, utcnow()))
        connection.execute("INSERT INTO audit_log(user_id,action,entity,details,created_at) VALUES(?,?,?,?,?)",
            (user_id, "save", "vat_return", json.dumps({"year": int(year), "quarter": int(quarter), "payable_lbp": str(result["payable_lbp"]),
             "credit_carried_forward_lbp": str(result["credit_carried_forward_lbp"])}), utcnow()))
    return build_vat_return(db, year, quarter, None, False, previous_year_db, credit_brought_forward)


def reopen_return(db, year, quarter, user_id):
    if not saved_return(db, year, quarter): raise ValueError("This VAT return has not been saved")
    with db.connect() as connection:
        connection.execute("DELETE FROM vat_returns WHERE year=? AND quarter=?", (int(year), int(quarter)))
        connection.execute("INSERT INTO audit_log(user_id,action,entity,details,created_at) VALUES(?,?,?,?,?)",
            (user_id, "reopen", "vat_return", json.dumps({"year": int(year), "quarter": int(quarter)}), utcnow()))
    return {"reopened": True}


def export_sections(result):
    """Turn a VAT return into titled table sections for Excel / PDF export."""
    title = f"Quarterly VAT Return - Q{result['quarter']} {result['year']}"
    meta = [f"Period: {display_date(result['date_from'])} to {display_date(result['date_to'])}   Currency filter: {result['currency_filter']}",
            f"Status: {result['status']}   Source: {'posted and review documents' if result['include_review'] else 'posted documents only'}"]
    if result["review_excluded"]: meta.append(f"Note: {result['review_excluded']} document(s) in Review status are excluded from this return")
    sections = []
    for code, values in sorted(result["per_currency"].items()):
        rows = [[number, label, values[key]["base"] if key not in ("adj_output", "adj_input", "net") else "",
                 values[key]["vat"], values[key]["vat_lbp"]] for number, key, label in LINES]
        sections.append({"heading": f"Return by currency - {code}", "headers": ["Line", "Description", f"Base ({code})", f"VAT ({code})", "VAT (LBP)"],
                         "rows": rows, "total_rows": [2, 8, 10]})
    totals = result["totals_lbp"]
    summary = [[number, label, totals[key]] for number, key, label in LINES]
    summary += [["12", f"Credit brought forward ({result['credit_source']})", result["credit_brought_forward_lbp"]],
                ["13", "VAT payable to the Ministry of Finance", result["payable_lbp"]],
                ["14", "Credit carried forward to next quarter", result["credit_carried_forward_lbp"]]]
    sections.append({"heading": "Total return - all currencies in LBP", "headers": ["Line", "Description", "Amount (LBP)"], "rows": summary, "total_rows": [2, 8, 10, 12, 13]})
    detail = [[display_date(d["date"]), d["number"], d["party"], CATEGORIES[d["category"]], "Yes" if d["recoverable"] else "No", d["currency"],
               d["base"], d["vat"], d["lbp_rate"], d["vat_lbp"], d["status"]] for d in result["documents"]]
    sections.append({"heading": "Supporting documents", "headers": ["Date", "Document", "Customer / Supplier", "Category", "Deductible", "Currency",
                     "Base", "VAT", "LBP Rate", "VAT (LBP)", "Status"], "rows": detail or [["No documents in this quarter"] + [""] * 10], "total_rows": []})
    if result["adjustments"]:
        sections.append({"heading": "Manual adjustments", "headers": ["Type", "Currency", "Amount", "Amount (LBP)", "Reason", "Entered by", "Entered at"],
            "rows": [[ADJUSTMENT_TYPES[a["adjustment_type"]], a["currency"], _money(a["amount"]), a.get("amount_lbp", ""), a["reason"],
                      a.get("created_by_name") or "", str(a["created_at"])[:16].replace("T", " ")] for a in result["adjustments"]], "total_rows": []})
    return title, meta, sections


def json_ready(value):
    if isinstance(value, Decimal): return float(value)
    if isinstance(value, list): return [json_ready(v) for v in value]
    if isinstance(value, dict): return {k: json_ready(v) for k, v in value.items()}
    return value
