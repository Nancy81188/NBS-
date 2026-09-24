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


def _header_map(sheet, keywords, scan=10):
    """Find the heading row and map each wanted field to its column index."""
    for row_number, row in enumerate(sheet.iter_rows(min_row=1, max_row=scan, values_only=True), 1):
        labels = [str(value or "").strip().casefold() for value in row]
        mapping = {}
        for field, words in keywords.items():
            for index, label in enumerate(labels):
                if label and any(word in label for word in words) and index not in mapping.values(): mapping[field] = index; break
        if len(mapping) >= 2: return row_number, mapping
    raise ValueError("The Excel file has no heading row that Saber recognises")


def read_expenses(path):
    """Expenses from Excel: Date, Description, Category, Currency, Amount (with VAT base), Without VAT, VAT, Reference."""
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = workbook.worksheets[0]
        header, columns = _header_map(sheet, {"date": ("date", "تاريخ"), "description": ("description", "details", "بيان", "libell"), "category": ("category", "type"),
            "currency": ("currency", "devise", "عملة"), "without_vat": ("without vat", "no vat", "exempt"), "vat": ("vat", "tva", "tax"),
            "amount": ("amount", "before vat", "subtotal", "montant", "مبلغ"), "reference": ("reference", "ref", "invoice", "رقم")})
        if "amount" not in columns and "without_vat" not in columns: raise ValueError("Add an Amount column to the Excel file")
        rows = []
        for number, row in enumerate(sheet.iter_rows(min_row=header + 1, values_only=True), header + 1):
            if not any(value not in (None, "") for value in row): continue
            get = lambda field: row[columns[field]] if field in columns and columns[field] < len(row) else None
            def money(field):
                value = get(field)
                try: return float(str(value).replace(",", "")) if value not in (None, "") else 0.0
                except ValueError: raise ValueError(f"Row {number}: {field.replace('_', ' ')} must be a number")
            amount = money("amount"); without = money("without_vat"); vat = money("vat") if "vat" in columns else round(amount * 0.11, 2)
            if not amount and not without: continue
            rows.append({"expense_date": _date(get("date")), "description": str(get("description") or get("reference") or f"Expense row {number}").strip(),
                         "category": str(get("category") or "General").strip(), "currency": str(get("currency") or "USD").strip().upper()[:3],
                         "with_vat_subtotal": amount, "without_vat_subtotal": without, "vat": vat, "reference": str(get("reference") or "").strip(), "source_row": number})
        return rows
    finally: workbook.close()


def read_customs_costs(path):
    """Landed costs from a customs / broker Excel: totals of Freight, Insurance, Duties, Broker fees, Other, Import VAT."""
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = workbook.worksheets[0]
        header, columns = _header_map(sheet, {"freight": ("freight", "fret", "shipping"), "insurance": ("insurance", "assurance"),
            "customs_duties": ("dut", "customs", "douane", "جمرك"), "broker_fees": ("broker", "clearance", "transit", "مخلص"),
            "import_vat": ("vat", "tva"), "other_costs": ("other", "port", "storage", "handling"), "customs_declaration_no": ("declaration", "bayan", "بيان", "decl")})
        totals = {key: 0.0 for key in columns if key != "customs_declaration_no"}; declaration = ""
        for row in sheet.iter_rows(min_row=header + 1, values_only=True):
            for key, index in columns.items():
                value = row[index] if index < len(row) else None
                if value in (None, ""): continue
                if key == "customs_declaration_no": declaration = declaration or str(value).strip(); continue
                try: totals[key] += float(str(value).replace(",", ""))
                except ValueError: pass
        return {**{k: round(v, 2) for k, v in totals.items()}, "customs_declaration_no": declaration}
    finally: workbook.close()

