from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
import re

from openpyxl import load_workbook

VAT_RATE = Decimal("0.11")
SUPPORTED_CURRENCIES = ("USD", "EUR", "LBP", "AED")

ALIASES = {
    "invoice_number": {"invoice number", "invoice no", "invoice no.", "invoice #", "inv no", "رقم الفاتورة", "numero facture", "n facture"},
    "date": {"date", "invoice date", "تاريخ", "التاريخ", "date facture"},
    "party": {"supplier name", "customer name", "client name", "supplier", "customer", "client", "المورد", "العميل", "fournisseur"},
    "subtotal": {"total before vat", "before vat", "subtotal", "amount before vat", "قبل الضريبة", "hors tva", "ht"},
    "vat": {"vat", "tva", "ضريبة", "الضريبة"},
    "total": {"total after vat", "total", "grand total", "بعد الضريبة", "ttc"},
    "currency": {"currency", "curr", "العملة", "devise"},
    "kind": {"type", "invoice type", "نوع", "nature"},
    "supplier_account": {"supplier account number", "supplier account", "supplier account no", "supplier a/c", "حساب المورد", "compte fournisseur"},
    "vat_account": {"vat account number", "vat account", "vat account no", "vat a/c", "حساب الضريبة", "compte tva"},
    "expense_account": {"expense account number", "expense account", "expense account no", "expense a/c", "classification number", "حساب المصروف", "compte charge"},
}

_CURRENCY_PATTERNS = {
    "USD": (
        re.compile(r"\$", re.IGNORECASE),
        re.compile(r"(?<![A-Z])USD(?![A-Z])", re.IGNORECASE),
    ),
    "EUR": (
        re.compile(r"€", re.IGNORECASE),
        re.compile(r"(?<![A-Z])EUR(?![A-Z])", re.IGNORECASE),
    ),
    "LBP": (
        re.compile(r"(?<![A-Z])LBP(?![A-Z])", re.IGNORECASE),
        re.compile(r"(?<![A-Z])L\s*\.\s*L\s*\.?(?![A-Z])", re.IGNORECASE),
        re.compile(r"ل\s*\.\s*ل"),
    ),
    "AED": (
        re.compile(r"(?<![A-Z])AED(?![A-Z])", re.IGNORECASE),
        re.compile(r"د\s*\.\s*إ"),
    ),
}

def _norm(value):
    return re.sub(r"[^\w#]+", " ", str(value or "").strip().lower()).strip()

def _columns(header):
    result = {}
    for index, value in enumerate(header):
        normalized = _norm(value)
        for field, aliases in ALIASES.items():
            if normalized in aliases:
                result[field] = index
                break
    return result

def _detected_currencies(*values):
    found = set()
    for value in values:
        text = str(value or "")
        for currency, patterns in _CURRENCY_PATTERNS.items():
            if any(pattern.search(text) for pattern in patterns):
                found.add(currency)
    return found

def _currency(values, columns, default_currency, number_formats=()):
    """Detect currency from price cells first, then use other evidence."""
    default = str(default_currency or "USD").strip().upper()
    if default not in SUPPORTED_CURRENCIES:
        default = "USD"

    currency_cell = ""
    currency_index = columns.get("currency")
    if currency_index is not None and currency_index < len(values):
        currency_cell = str(values[currency_index] or "").strip()
    explicit = _detected_currencies(currency_cell)

    votes = {code: 0 for code in SUPPORTED_CURRENCIES}
    field_currencies = {}

    # The symbols/codes printed in monetary cells are the primary source.
    for field in ("subtotal", "vat", "total"):
        index = columns.get(field)
        if index is None or index >= len(values) or values[index] in (None, ""):
            continue
        evidence = [values[index]]
        if index < len(number_formats):
            evidence.append(number_formats[index])
        detected = _detected_currencies(*evidence)
        field_currencies[field] = detected
        for code in detected:
            votes[code] += 1

    found = {code for code, count in votes.items() if count}

    if found:
        highest = max(votes.values())
        candidates = {code for code, count in votes.items() if count == highest}
        selected = None

        # Prefer the final price, then subtotal, when votes are tied.
        for field in ("total", "subtotal", "vat"):
            detected = field_currencies.get(field, set()) & candidates
            if len(detected) == 1:
                selected = next(iter(detected))
                break

        if selected is None:
            non_default = [
                code for code in ("EUR", "LBP", "AED")
                if code in candidates
            ]
            selected = non_default[0] if non_default else "USD"

        # Currency formatting conflicts are resolved automatically so the row
        # appears only in the selected currency view.
        return selected, ""

    # Only consult a separate Currency cell when prices contain no evidence.
    if len(explicit) == 1:
        return next(iter(explicit)), ""
    if len(explicit) > 1:
        selected = next(
            (code for code in ("EUR", "LBP", "AED") if code in explicit),
            "USD",
        )
        return selected, ""
    if currency_cell:
        return default, "unsupported:" + currency_cell
    return default, "missing_defaulted_to_" + default.lower()

def _decimal(value):
    if value in (None, ""):
        return None
    if isinstance(value, str):
        negative = value.strip().startswith("(") and value.strip().endswith(")")
        value = re.sub(r"(?i)USD|EUR|LBP|AED", "", value)
        value = re.sub(r"L\s*\.\s*L\s*\.?", "", value, flags=re.IGNORECASE)
        value = re.sub(r"[ل]\s*\.\s*[ل]|[د]\s*\.\s*[إ]", "", value)
        value = value.replace("$", "").replace("€", "").replace(",", "").replace(" ", "").strip("() ")
        if negative:
            value = "-" + value
    try:
        return Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except (InvalidOperation, ValueError):
        return None

def _date(value):
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    text = str(value or "").strip()
    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d", "%m/%d/%Y"):
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            pass
    return text

def read_invoices(path: str | Path, sheet_name: str | None = None, default_currency="USD", default_kind="purchase"):
    """Read invoices without deleting duplicates; completely blank rows are ignored.

    Currency is detected from the Currency and monetary columns. Missing currency
    uses the selected default (USD by default); conflicts and unsupported values
    are returned in currency_issue for review.
    """
    workbook_values = load_workbook(path, read_only=True, data_only=True)
    workbook_formulas = load_workbook(path, read_only=True, data_only=False)
    try:
        sheet = workbook_values[sheet_name] if sheet_name else workbook_values.worksheets[0]
        formula_sheet = workbook_formulas[sheet.title]
        value_rows = sheet.iter_rows(values_only=True)
        formula_rows = formula_sheet.iter_rows()
        header = next(value_rows)
        next(formula_rows)
        columns = _columns(header)
        required = {"date", "party", "subtotal", "vat", "total"}
        if not required.issubset(columns):
            missing = ", ".join(sorted(required - set(columns)))
            raise ValueError(f"Missing required columns: {missing}")
        invoices = []
        for row_number, (values, formula_cells) in enumerate(zip(value_rows, formula_rows), start=2):
            if not any(cell.value not in (None, "") for cell in formula_cells):
                continue
            number_formats = tuple(cell.number_format or "" for cell in formula_cells)
            def get(field, default=None):
                idx = columns.get(field)
                return values[idx] if idx is not None and idx < len(values) else default
            subtotal, vat, total = _decimal(get("subtotal")), _decimal(get("vat")), _decimal(get("total"))
            if subtotal is not None and vat is None:
                vat = (subtotal * VAT_RATE).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            if total is None and subtotal is not None and vat is not None:
                total = subtotal + vat
            invoice_number = str(get("invoice_number") or row_number).strip()
            kind = str(get("kind") or default_kind).strip().lower()
            currency, currency_issue = _currency(values, columns, default_currency, number_formats)
            invoices.append({
                "invoice_number": invoice_number,
                "invoice_date": _date(get("date")),
                "party_name": str(get("party") or "").strip(),
                "subtotal": float(subtotal) if subtotal is not None else None,
                "vat": float(vat) if vat is not None else None,
                "total": float(total) if total is not None else None,
                "currency": currency,
                "currency_issue": currency_issue,
                "kind": "sale" if kind in {"sale", "sales", "customer"} else "purchase",
                "supplier_account": str(get("supplier_account") or "4011").strip(),
                "vat_account": str(get("vat_account") or "442660000").strip(),
                "expense_account": str(get("expense_account") or "601100000").strip(),
                "source_file": Path(path).name,
                "source_row": row_number,
            })
        return invoices
    finally:
        workbook_values.close()
        workbook_formulas.close()
