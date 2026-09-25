"""Inventory: items, warehouses, stock documents, weighted-average / FIFO costing, reports and the
year-end stock variation voucher of the Lebanese chart (6051 opening stock / 6052 closing stock / 37 stock).

Accounting follows the periodic method of the Lebanese chart of accounts: purchases stay in 601, and the
stock value is booked at the period end by a 'STOCK VARIATION' Journal Voucher (type 06)."""
from __future__ import annotations

import json
from datetime import datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP

from database import display_date, iso_date, utcnow

ZERO = Decimal("0")
DOC_TYPES = {"opening": ("OPN", "Opening Stock", 1), "receipt": ("GRN", "Stock Receipt", 1), "issue": ("GIN", "Stock Issue", -1),
             "adjustment_in": ("ADJ", "Adjustment +", 1), "adjustment_out": ("ADJ", "Adjustment -", -1), "transfer": ("TRF", "Transfer", 0)}
STOCK_ACCOUNT, OPENING_ACCOUNT, CLOSING_ACCOUNT = "37", "6051", "6052"


def _d(value):
    try: return Decimal(str(value if value not in (None, "") else 0).replace(",", ""))
    except Exception as exc: raise ValueError(f"'{value}' is not a number") from exc


def migrate(db):
    """Called from Database.initialize: inventory tables and columns."""
    db.execute("""CREATE TABLE IF NOT EXISTS warehouses (id INTEGER PRIMARY KEY, code TEXT NOT NULL UNIQUE, name TEXT NOT NULL, active INTEGER NOT NULL DEFAULT 1)""")
    db.execute("""CREATE TABLE IF NOT EXISTS stock_documents (
        id INTEGER PRIMARY KEY, number TEXT NOT NULL UNIQUE, doc_type TEXT NOT NULL, doc_date TEXT NOT NULL, warehouse_id INTEGER NOT NULL,
        to_warehouse_id INTEGER, party_id INTEGER, invoice_id INTEGER, reference TEXT, notes TEXT, created_by INTEGER, created_at TEXT NOT NULL)""")
    item_columns = {row["name"] for row in db.execute("PRAGMA table_info(inventory_items)")}
    for column, definition in (("category", "TEXT"), ("reorder_level", "TEXT NOT NULL DEFAULT '0'"), ("sales_price", "TEXT NOT NULL DEFAULT '0'"),
                               ("active", "INTEGER NOT NULL DEFAULT 1"), ("notes", "TEXT"), ("barcode", "TEXT"), ("created_at", "TEXT")):
        if column not in item_columns: db.execute(f"ALTER TABLE inventory_items ADD COLUMN {column} {definition}")
    movement_columns = {row["name"] for row in db.execute("PRAGMA table_info(stock_movements)")}
    for column, definition in (("warehouse_id", "INTEGER"), ("document_id", "INTEGER"), ("movement_type", "TEXT"), ("sales_price", "TEXT"), ("line_no", "INTEGER")):
        if column not in movement_columns: db.execute(f"ALTER TABLE stock_movements ADD COLUMN {column} {definition}")
    db.execute("CREATE TABLE IF NOT EXISTS item_categories (id INTEGER PRIMARY KEY, name TEXT NOT NULL, parent_id INTEGER, UNIQUE(name,parent_id))")
    db.execute("CREATE TABLE IF NOT EXISTS item_units (id INTEGER PRIMARY KEY, name TEXT NOT NULL UNIQUE)")
    db.execute("""CREATE TABLE IF NOT EXISTS physical_counts (id INTEGER PRIMARY KEY, number TEXT NOT NULL UNIQUE, count_date TEXT NOT NULL, warehouse_id INTEGER NOT NULL,
        status TEXT NOT NULL DEFAULT 'draft', lines TEXT NOT NULL, adjustment_numbers TEXT, notes TEXT, created_by INTEGER, created_at TEXT NOT NULL)""")
    item_columns = {row["name"] for row in db.execute("PRAGMA table_info(inventory_items)")}
    for column in ("subcategory", "supplier_id", "location"):
        if column not in item_columns: db.execute(f"ALTER TABLE inventory_items ADD COLUMN {column} TEXT")
    for unit in ("unit", "piece", "sheet", "m", "m2", "kg", "box", "roll", "set"): db.execute("INSERT OR IGNORE INTO item_units(name) VALUES(?)", (unit,))
    db.execute("INSERT OR IGNORE INTO warehouses(code,name) VALUES('MAIN','Main Store')")
    db.execute("INSERT OR IGNORE INTO app_settings(key,value) VALUES('inventory_currency','USD')")
    db.execute("INSERT OR IGNORE INTO app_settings(key,value) VALUES('inventory_method','average')")
    db.execute("UPDATE stock_movements SET warehouse_id=(SELECT id FROM warehouses WHERE code='MAIN') WHERE warehouse_id IS NULL")


def settings(database):
    values = database.settings()
    return {"currency": values.get("inventory_currency", "USD"), "method": values.get("inventory_method", "average")}


def save_settings(database, item, user_id):
    currency = str(item.get("currency") or "USD").upper(); method = str(item.get("method") or "average").lower()
    if currency not in ("USD", "LBP", "EUR", "AED") or method not in ("average", "fifo"): raise ValueError("Choose USD/LBP/EUR/AED and Average or FIFO")
    with database.connect() as db:
        for key, value in (("inventory_currency", currency), ("inventory_method", method)):
            db.execute("INSERT INTO app_settings(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value", (key, value))
    return settings(database)


# ---------------------------------------------------------------- master data
def list_warehouses(database):
    with database.connect() as db: return [dict(r) for r in db.execute("SELECT * FROM warehouses ORDER BY code")]


def save_warehouse(database, item, user_id):
    name = str(item.get("name") or "").strip(); code = str(item.get("code") or "").strip().upper()
    if not name: raise ValueError("Warehouse name is required")
    with database.connect() as db:
        if not code:
            numbers = [int(r["code"][2:]) for r in db.execute("SELECT code FROM warehouses WHERE code GLOB 'WH[0-9]*'") if r["code"][2:].isdigit()]
            code = f"WH{max(numbers, default=0) + 1:02d}"
        if db.execute("SELECT 1 FROM warehouses WHERE code=? AND id<>?", (code, int(item.get("id") or 0))).fetchone(): raise ValueError(f"Warehouse code {code} is already used")
        if item.get("id"): db.execute("UPDATE warehouses SET code=?,name=?,active=? WHERE id=?", (code, name, 1 if item.get("active", True) else 0, int(item["id"]))); saved = int(item["id"])
        else: saved = db.execute("INSERT INTO warehouses(code,name,active) VALUES(?,?,?)", (code, name, 1 if item.get("active", True) else 0)).lastrowid
    return next(w for w in list_warehouses(database) if w["id"] == saved)


def save_item(database, item, user_id):
    name = str(item.get("name") or "").strip(); sku = str(item.get("sku") or "").strip().upper()
    if not name: raise ValueError("Item name is required")
    reorder = _d(item.get("reorder_level")); price = _d(item.get("sales_price"))
    if reorder < 0 or price < 0: raise ValueError("Reorder level and sales price cannot be negative")
    with database.connect() as db:
        if not sku:
            numbers = [int(r["sku"][4:]) for r in db.execute("SELECT sku FROM inventory_items WHERE sku GLOB 'ITM-[0-9]*'") if r["sku"][4:].isdigit()]
            sku = f"ITM-{max(numbers, default=0) + 1:05d}"
        if db.execute("SELECT 1 FROM inventory_items WHERE sku=? AND id<>?", (sku, int(item.get("id") or 0))).fetchone(): raise ValueError(f"Item code {sku} is already used")
        supplier = item.get("supplier_id")
        if not supplier and str(item.get("supplier_name") or "").strip():
            row = db.execute("SELECT id FROM parties WHERE name=? ORDER BY id LIMIT 1", (str(item["supplier_name"]).strip(),)).fetchone(); supplier = row["id"] if row else None
        unit = str(item.get("unit") or "unit").strip() or "unit"; db.execute("INSERT OR IGNORE INTO item_units(name) VALUES(?)", (unit,))
        values = (sku, name, unit, str(item.get("category") or "").strip() or None, str(reorder), str(price),
                  1 if item.get("active", True) else 0, str(item.get("notes") or "").strip() or None, str(item.get("barcode") or "").strip() or None,
                  str(item.get("subcategory") or "").strip() or None, str(supplier) if supplier else None, str(item.get("location") or "").strip() or None)
        if item.get("id"):
            db.execute("UPDATE inventory_items SET sku=?,name=?,unit=?,category=?,reorder_level=?,sales_price=?,active=?,notes=?,barcode=?,subcategory=?,supplier_id=?,location=? WHERE id=?", values + (int(item["id"]),)); saved = int(item["id"])
        else:
            saved = db.execute("INSERT INTO inventory_items(sku,name,unit,category,reorder_level,sales_price,active,notes,barcode,subcategory,supplier_id,location,created_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)", values + (utcnow(),)).lastrowid
        db.execute("INSERT INTO audit_log(user_id,action,entity,entity_id,details,created_at) VALUES(?,?,?,?,?,?)", (user_id, "save", "inventory_item", saved, json.dumps({"sku": sku}), utcnow()))
    return next(i for i in list_items(database) if i["id"] == saved)


# ---------------------------------------------------------------- costing engine
def _movements(database, date_to=None, exclude_document_id=None):
    with database.connect() as db:
        rows = [dict(r) for r in db.execute("""SELECT m.*,d.number,d.doc_type,d.doc_date,d.reference,d.party_id,d.invoice_id,p.name party_name,w.code warehouse_code
            FROM stock_movements m JOIN stock_documents d ON d.id=m.document_id LEFT JOIN parties p ON p.id=d.party_id LEFT JOIN warehouses w ON w.id=m.warehouse_id
            ORDER BY d.doc_date,d.id,m.id""")]
    return [r for r in rows if (not date_to or r["doc_date"] <= date_to) and r["document_id"] != exclude_document_id]


def run_costing(database, date_to=None, method=None, callback=None, exclude_document_id=None):
    """Replays every movement in date order. Returns per-item {qty, value, by_warehouse, last_date}.

    Average: each receipt re-weights the cost; issues leave at the running average.
    FIFO: issues consume the oldest receipt layers first. Transfers move quantity only."""
    method = method or settings(database)["method"]; state = {}
    for row in _movements(database, date_to, exclude_document_id):
        item = state.setdefault(row["item_id"], {"qty": ZERO, "value": ZERO, "layers": [], "by_warehouse": {}, "last_date": None, "last_out": None})
        qty = _d(row["quantity"]); cost = _d(row["unit_cost"])
        item["by_warehouse"][row["warehouse_id"]] = item["by_warehouse"].get(row["warehouse_id"], ZERO) + qty
        item["last_date"] = row["doc_date"]
        if row["doc_type"] == "transfer":
            unit = (item["value"] / item["qty"]) if item["qty"] else ZERO
            if callback: callback(row, unit, ZERO)
            continue
        if qty > 0:
            item["qty"] += qty; item["value"] += qty * cost; item["layers"].append([qty, cost]); issued_cost = cost
        else:
            out = -qty
            if method == "fifo":
                remaining = out; total = ZERO
                while remaining > 0 and item["layers"]:
                    layer = item["layers"][0]; take = min(layer[0], remaining); total += take * layer[1]; layer[0] -= take; remaining -= take
                    if layer[0] <= 0: item["layers"].pop(0)
                if remaining > 0: total += remaining * ((item["value"] / item["qty"]) if item["qty"] else ZERO)
                issued_cost = total / out if out else ZERO
            else:
                issued_cost = (item["value"] / item["qty"]) if item["qty"] else ZERO
                if method != "fifo":
                    for layer in item["layers"]: layer[1] = issued_cost
            item["qty"] -= out; item["value"] -= out * issued_cost; item["last_out"] = row["doc_date"]
            if item["qty"] <= 0: item["value"] = ZERO; item["layers"] = []
        if callback: callback(row, issued_cost, qty * issued_cost)
    for item in state.values():
        item["avg"] = (item["value"] / item["qty"]) if item["qty"] > 0 else ZERO
    return state


def list_items(database, date_to=None, include_inactive=True):
    state = run_costing(database, iso_date(date_to) if date_to else None)
    with database.connect() as db:
        items = [dict(r) for r in db.execute("SELECT i.*,p.name supplier_name FROM inventory_items i LEFT JOIN parties p ON p.id=CAST(i.supplier_id AS INTEGER) ORDER BY i.sku")]
        last = {r["item_id"]: r["party_name"] for r in db.execute("""SELECT m.item_id,p.name party_name FROM stock_movements m JOIN stock_documents d ON d.id=m.document_id
            JOIN parties p ON p.id=d.party_id WHERE d.doc_type='receipt' ORDER BY d.doc_date,d.id""")}
    for item in items:
        item["supplier_name"] = item.get("supplier_name") or last.get(item["id"]) or ""
        data = state.get(item["id"], {})
        item["quantity"] = float(data.get("qty", ZERO)); item["average_cost"] = float(data.get("avg", ZERO)); item["stock_value"] = float(data.get("value", ZERO))
        item["reorder_level"] = float(_d(item.get("reorder_level"))); item["sales_price"] = float(_d(item.get("sales_price")))
    return [i for i in items if include_inactive or i["active"]]


# ---------------------------------------------------------------- stock documents
def next_number(database, doc_type, date):
    prefix = DOC_TYPES[doc_type][0]; year = iso_date(date)[:4]
    with database.connect() as db:
        numbers = [int(r["number"].rsplit("-", 1)[-1]) for r in db.execute("SELECT number FROM stock_documents WHERE number LIKE ?", (f"{prefix}-{year}-%",)) if r["number"].rsplit("-", 1)[-1].isdigit()]
    return f"{prefix}-{year}-{max(numbers, default=0) + 1:06d}"


def save_document(database, header, lines, user_id, document_id=None):
    doc_type = str(header.get("doc_type") or "").lower()
    if doc_type not in DOC_TYPES: raise ValueError("Choose the document type")
    date = iso_date(header.get("doc_date"), "Date"); database._assert_period_open(date)
    if not isinstance(lines, list) or not lines: raise ValueError("Add at least one item line")
    with database.connect() as db:
        warehouse = db.execute("SELECT * FROM warehouses WHERE id=? OR code=?", (int(header["warehouse_id"]) if str(header.get("warehouse_id") or "").isdigit() else -1, str(header.get("warehouse_id") or "MAIN"))).fetchone()
        if not warehouse: raise ValueError("Choose the warehouse")
        target = None
        if doc_type == "transfer":
            target = db.execute("SELECT * FROM warehouses WHERE id=? OR code=?", (int(header["to_warehouse_id"]) if str(header.get("to_warehouse_id") or "").isdigit() else -1, str(header.get("to_warehouse_id") or ""))).fetchone()
            if not target or target["id"] == warehouse["id"]: raise ValueError("Choose a different destination warehouse for the transfer")
        items = {}; normalized = []
        for index, line in enumerate(lines, 1):
            code = str(line.get("sku") or "").strip().upper()
            item = db.execute("SELECT * FROM inventory_items WHERE sku=? OR id=?", (code, int(line["item_id"]) if str(line.get("item_id") or "").isdigit() else -1)).fetchone()
            if not item: raise ValueError(f"Line {index}: item {code or line.get('item_id')} was not found")
            qty = _d(line.get("quantity")); cost = _d(line.get("unit_cost"))
            if qty <= 0: raise ValueError(f"Line {index}: quantity must be above zero")
            if cost < 0: raise ValueError(f"Line {index}: unit cost cannot be negative")
            if DOC_TYPES[doc_type][2] > 0 and not cost and doc_type != "adjustment_in": raise ValueError(f"Line {index}: enter the unit cost of {item['sku']}")
            normalized.append((item, qty, cost, _d(line.get("sales_price")))); items[item["id"]] = items.get(item["id"], ZERO) + qty
    # Stock may not go negative: check the quantity available in the warehouse on the document date.
    if DOC_TYPES[doc_type][2] <= 0:
        state = run_costing(database, date) if not document_id else _state_without(database, date, document_id)
        for item_id, qty in items.items():
            available = state.get(item_id, {}).get("by_warehouse", {}).get(warehouse["id"], ZERO)
            if qty > available:
                sku = next(n[0]["sku"] for n in normalized if n[0]["id"] == item_id)
                raise ValueError(f"Not enough stock of {sku} in {warehouse['code']} on {display_date(date)}: available {available:,.3f}, requested {qty:,.3f}")
    with database.connect() as db:
        if document_id:
            old = db.execute("SELECT * FROM stock_documents WHERE id=?", (int(document_id),)).fetchone()
            if not old: raise KeyError("Stock document not found")
            database._assert_period_open(old["doc_date"]); number = old["number"]
            db.execute("DELETE FROM stock_movements WHERE document_id=?", (int(document_id),))
            db.execute("""UPDATE stock_documents SET doc_type=?,doc_date=?,warehouse_id=?,to_warehouse_id=?,party_id=?,reference=?,notes=? WHERE id=?""",
                (doc_type, date, warehouse["id"], target["id"] if target else None, header.get("party_id") or None, header.get("reference") or None, header.get("notes") or None, int(document_id)))
            saved = int(document_id)
        else:
            number = str(header.get("number") or "").strip() or next_number(database, doc_type, date)
            saved = db.execute("""INSERT INTO stock_documents(number,doc_type,doc_date,warehouse_id,to_warehouse_id,party_id,invoice_id,reference,notes,created_by,created_at)
                VALUES(?,?,?,?,?,?,?,?,?,?,?)""", (number, doc_type, date, warehouse["id"], target["id"] if target else None, header.get("party_id") or None,
                header.get("invoice_id") or None, header.get("reference") or None, header.get("notes") or None, user_id, utcnow())).lastrowid
        sign = DOC_TYPES[doc_type][2]
        for position, (item, qty, cost, price) in enumerate(normalized, 1):
            if doc_type == "transfer":
                for warehouse_id, signed in ((warehouse["id"], -qty), (target["id"], qty)):
                    db.execute("""INSERT INTO stock_movements(item_id,movement_date,quantity,unit_cost,source_type,source_id,warehouse_id,document_id,movement_type,line_no)
                        VALUES(?,?,?,?,?,?,?,?,?,?)""", (item["id"], date, str(signed), "0", "stock", saved, warehouse_id, saved, doc_type, position))
            else:
                db.execute("""INSERT INTO stock_movements(item_id,movement_date,quantity,unit_cost,source_type,source_id,warehouse_id,document_id,movement_type,sales_price,line_no)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?)""", (item["id"], date, str(qty * sign), str(cost), "stock", saved, warehouse["id"], saved, doc_type, str(price), position))
        db.execute("INSERT INTO audit_log(user_id,action,entity,entity_id,details,created_at) VALUES(?,?,?,?,?,?)",
            (user_id, "save", "stock_document", saved, json.dumps({"number": number, "type": doc_type, "lines": len(normalized)}), utcnow()))
    return get_document(database, saved)


def _state_without(database, date, document_id):
    return run_costing(database, date, exclude_document_id=int(document_id))


def get_document(database, document_id):
    with database.connect() as db:
        doc = db.execute("""SELECT d.*,w.code warehouse_code,w.name warehouse_name,t.code to_warehouse_code,p.name party_name FROM stock_documents d
            JOIN warehouses w ON w.id=d.warehouse_id LEFT JOIN warehouses t ON t.id=d.to_warehouse_id LEFT JOIN parties p ON p.id=d.party_id WHERE d.id=?""", (int(document_id),)).fetchone()
        if not doc: raise KeyError("Stock document not found")
        lines = [dict(r) for r in db.execute("""SELECT m.*,i.sku,i.name,i.unit FROM stock_movements m JOIN inventory_items i ON i.id=m.item_id
            WHERE m.document_id=? ORDER BY m.line_no,m.id""", (int(document_id),))]
    doc = dict(doc)
    if doc["doc_type"] == "transfer": lines = [l for l in lines if _d(l["quantity"]) > 0]
    for line in lines: line["quantity"] = float(abs(_d(line["quantity"]))); line["unit_cost"] = float(_d(line["unit_cost"])); line["sales_price"] = float(_d(line.get("sales_price")))
    doc["lines"] = lines
    return doc


def list_documents(database):
    with database.connect() as db:
        return [dict(r) for r in db.execute("""SELECT d.id,d.number,d.doc_type,d.doc_date,w.code warehouse_code,t.code to_warehouse_code,p.name party_name,d.reference,d.invoice_id,
            (SELECT COUNT(DISTINCT line_no) FROM stock_movements m WHERE m.document_id=d.id) lines FROM stock_documents d JOIN warehouses w ON w.id=d.warehouse_id
            LEFT JOIN warehouses t ON t.id=d.to_warehouse_id LEFT JOIN parties p ON p.id=d.party_id ORDER BY d.doc_date DESC,d.id DESC""")]


def delete_document(database, document_id, user_id):
    doc = get_document(database, document_id); database._assert_period_open(doc["doc_date"])
    if DOC_TYPES[doc["doc_type"]][2] > 0:  # removing stock that was already issued afterwards would make it negative
        _check_after_removal(database, document_id)
    with database.connect() as db:
        db.execute("DELETE FROM stock_movements WHERE document_id=?", (int(document_id),)); db.execute("DELETE FROM stock_documents WHERE id=?", (int(document_id),))
        db.execute("INSERT INTO audit_log(user_id,action,entity,entity_id,details,created_at) VALUES(?,?,?,?,?,?)", (user_id, "delete", "stock_document", int(document_id), json.dumps({"number": doc["number"]}), utcnow()))
    return {"deleted": int(document_id)}


def _check_after_removal(database, document_id):
    balances = {}
    for row in _movements(database, exclude_document_id=int(document_id)):
        key = (row["item_id"], row["warehouse_id"]); balances[key] = balances.get(key, ZERO) + _d(row["quantity"])
        if balances[key] < 0: raise ValueError(f"This document cannot be deleted: the stock issued later ({row['number']} on {display_date(row['doc_date'])}) would become negative")


def issue_for_invoice(database, invoice_id, lines, user_id):
    """Sales invoice lines that name an item leave the stock (Stock Issue linked to the invoice)."""
    with database.connect() as db:
        invoice = db.execute("SELECT * FROM invoices WHERE id=?", (int(invoice_id),)).fetchone()
        for old in db.execute("SELECT id FROM stock_documents WHERE invoice_id=?", (int(invoice_id),)).fetchall():
            db.execute("DELETE FROM stock_movements WHERE document_id=?", (old["id"],)); db.execute("DELETE FROM stock_documents WHERE id=?", (old["id"],))
    stock_lines = [{"sku": l.get("item_code") or l.get("sku"), "quantity": l.get("quantity"), "sales_price": l.get("unit_price")} for l in lines if l.get("item_code") or l.get("sku")]
    if not stock_lines or not invoice: return None
    header = {"doc_type": "issue" if invoice["kind"] == "sale" else "receipt", "doc_date": invoice["invoice_date"], "warehouse_id": lines[0].get("warehouse") or "MAIN",
              "party_id": invoice["party_id"], "invoice_id": int(invoice_id), "reference": invoice["invoice_number"], "notes": f"Invoice {invoice['invoice_number']}"}
    if invoice["kind"] != "sale":
        for line, source in zip(stock_lines, [l for l in lines if l.get("item_code") or l.get("sku")]): line["unit_cost"] = source.get("unit_price")
    return save_document(database, header, stock_lines, user_id)


def remove_invoice_documents(db, invoice_id):
    for old in db.execute("SELECT id FROM stock_documents WHERE invoice_id=?", (int(invoice_id),)).fetchall():
        db.execute("DELETE FROM stock_movements WHERE document_id=?", (old["id"],)); db.execute("DELETE FROM stock_documents WHERE id=?", (old["id"],))


# ---------------------------------------------------------------- reports
def _names(database):
    with database.connect() as db:
        items = {r["id"]: dict(r) for r in db.execute("SELECT * FROM inventory_items")}
        warehouses = {r["id"]: dict(r) for r in db.execute("SELECT * FROM warehouses")}
    return items, warehouses


def build_report(database, report, options):
    options = dict(options or {}); inv = settings(database); method = options.get("method") or inv["method"]; currency = inv["currency"]
    date_to = iso_date(options["date_to"]) if options.get("date_to") else datetime.now().strftime("%Y-%m-%d")
    date_from = iso_date(options["date_from"]) if options.get("date_from") else f"{date_to[:4]}-01-01"
    warehouse = options.get("warehouse_id"); warehouse = int(warehouse) if str(warehouse or "").isdigit() else None
    items, warehouses = _names(database); company = database.settings(); sections = []
    filters = [f"{label}: {options[key]}" for key, label in (("category", "Category"), ("subcategory", "Subcategory"), ("unit", "Unit"), ("supplier_name", "Supplier")) if options.get(key)]
    wanted = lambda item_id: not options.get("item_id") or int(options["item_id"]) == item_id
    category = str(options.get("category") or "").strip(); subcategory = str(options.get("subcategory") or "").strip()
    unit_filter = str(options.get("unit") or "").strip(); supplier_filter = str(options.get("supplier_id") or "").strip()
    listed = {i["id"]: i for i in list_items(database)} if supplier_filter else {}
    def in_category(item_id):
        item = items[item_id]
        if category and (item.get("category") or "") != category: return False
        if subcategory and (item.get("subcategory") or "") != subcategory: return False
        if unit_filter and (item.get("unit") or "") != unit_filter: return False
        if supplier_filter:
            data = listed.get(item_id, {})
            if str(item.get("supplier_id") or "") != supplier_filter and data.get("supplier_name") != options.get("supplier_name"): return False
        return True
    if report == "valuation":
        state = run_costing(database, date_to, method)
        headers = ["Item Code", "Item", "Category", "Unit", "Quantity", f"Unit Cost ({currency})", f"Stock Value ({currency})", "Sales Price", "Value at Sales Price", "Reorder Level", "Status"]
        rows = []; total = ZERO; sales_total = ZERO
        for item_id, item in sorted(items.items(), key=lambda pair: pair[1]["sku"]):
            if not in_category(item_id): continue
            data = state.get(item_id, {"qty": ZERO, "avg": ZERO, "by_warehouse": {}})
            qty = data["by_warehouse"].get(warehouse, ZERO) if warehouse else data["qty"]
            if not qty and not options.get("include_zero"): continue
            value = (qty * data["avg"]).quantize(Decimal("0.01")); price = _d(item.get("sales_price")); reorder = _d(item.get("reorder_level"))
            total += value; sales_total += qty * price
            rows.append([item["sku"], item["name"], item.get("category") or "", item["unit"], qty, data["avg"].quantize(Decimal("0.0001")), value, price, (qty * price).quantize(Decimal("0.01")), reorder,
                         "Reorder" if reorder and qty <= reorder else "OK"])
        rows.append(["TOTAL", f"{len(rows)} item(s)", "", "", "", "", total, "", sales_total.quantize(Decimal("0.01")), "", ""])
        title = "Stock Valuation"; sections.append({"heading": f"Stock valuation at {display_date(date_to)} - {'weighted average' if method != 'fifo' else 'FIFO'}" + (f" - {warehouses[warehouse]['code']}" if warehouse else " - all warehouses"),
                                                     "headers": headers, "rows": rows, "total_rows": [len(rows) - 1]})
        if not warehouse and len(warehouses) > 1:
            by_wh = [[w["code"], w["name"], sum((data["by_warehouse"].get(wid, ZERO) * data["avg"] for data in state.values()), ZERO).quantize(Decimal("0.01"))] for wid, w in warehouses.items()]
            sections.append({"heading": "Value by warehouse", "headers": ["Warehouse", "Name", f"Value ({currency})"], "rows": by_wh, "total_rows": []})
    elif report == "stock_card":
        if not options.get("item_id"): raise ValueError("Choose the item for the stock card")
        item_id = int(options["item_id"]); lines = []; opening = {"qty": ZERO, "value": ZERO}
        def record(row, unit, value):
            if row["item_id"] != item_id or (warehouse and row["warehouse_id"] != warehouse): return
            qty = _d(row["quantity"]); cost_value = qty * (_d(row["unit_cost"]) if qty > 0 and row["doc_type"] != "transfer" else unit)
            if row["doc_date"] < date_from: opening["qty"] += qty; opening["value"] += cost_value; return
            lines.append([display_date(row["doc_date"]), row["number"], DOC_TYPES[row["doc_type"]][1], row.get("warehouse_code") or "", row.get("party_name") or row.get("reference") or "",
                          qty if qty > 0 else "", -qty if qty < 0 else "", (unit if qty < 0 or row["doc_type"] == "transfer" else _d(row["unit_cost"])).quantize(Decimal("0.0001")), cost_value.quantize(Decimal("0.01"))])
        run_costing(database, date_to, method, record)
        rows = [["", "", "Opening balance", "", "", "", "", "", opening["value"].quantize(Decimal("0.01"))]]; qty_balance = opening["qty"]; value_balance = opening["value"]
        rows[0].insert(9, qty_balance); rows[0].append(value_balance.quantize(Decimal("0.01")))
        total_in = ZERO; total_out = ZERO
        for line in lines:
            qty = _d(line[5] or 0) - _d(line[6] or 0); qty_balance += qty; value_balance += line[8]
            total_in += _d(line[5] or 0); total_out += _d(line[6] or 0)
            rows.append(line + [qty_balance, value_balance.quantize(Decimal("0.01"))])
        rows.append(["", "", "TOTAL / CLOSING", "", "", total_in, total_out, "", "", qty_balance, value_balance.quantize(Decimal("0.01"))])
        title = "Stock Card"
        item = items[item_id]
        sections.append({"heading": f"{item['sku']} - {item['name']} ({item['unit']})" + (f" - {warehouses[warehouse]['code']}" if warehouse else ""),
                         "headers": ["Date", "Document", "Type", "Warehouse", "Party / Reference", "In", "Out", f"Unit Cost ({currency})", "Value", "Balance Qty", "Balance Value"],
                         "rows": rows, "total_rows": [0, len(rows) - 1]})
    elif report == "movements":
        rows = []
        def record(row, unit, value):
            if row["doc_date"] < date_from or not wanted(row["item_id"]) or not in_category(row["item_id"]) or (warehouse and row["warehouse_id"] != warehouse): return
            if options.get("doc_type") and row["doc_type"] != options["doc_type"]: return
            qty = _d(row["quantity"]); unit_cost = _d(row["unit_cost"]) if qty > 0 and row["doc_type"] != "transfer" else unit
            rows.append([display_date(row["doc_date"]), row["number"], DOC_TYPES[row["doc_type"]][1], items[row["item_id"]]["sku"], items[row["item_id"]]["name"], row.get("warehouse_code") or "",
                         row.get("party_name") or "", qty, unit_cost.quantize(Decimal("0.0001")), (qty * unit_cost).quantize(Decimal("0.01"))])
        run_costing(database, date_to, method, record); title = "Stock Movements"
        sections.append({"heading": f"Movements {display_date(date_from)} to {display_date(date_to)}", "headers": ["Date", "Document", "Type", "Item Code", "Item", "Warehouse", "Party", "Quantity (+in / -out)", "Unit Cost", "Value"],
                         "rows": rows or [["No movements"] + [""] * 9], "total_rows": []})
    elif report == "margin":
        sold = {}
        def record(row, unit, value):
            if row["doc_type"] != "issue" or row["doc_date"] < date_from or not in_category(row["item_id"]): return
            data = sold.setdefault(row["item_id"], {"qty": ZERO, "cost": ZERO, "sales": ZERO})
            qty = -_d(row["quantity"]); data["qty"] += qty; data["cost"] += qty * unit; data["sales"] += qty * _d(row.get("sales_price"))
        run_costing(database, date_to, method, record)
        rows = []; totals = [ZERO, ZERO, ZERO]
        for item_id, data in sorted(sold.items(), key=lambda pair: -pair[1]["sales"]):
            margin = data["sales"] - data["cost"]; totals[0] += data["sales"]; totals[1] += data["cost"]; totals[2] += margin
            rows.append([items[item_id]["sku"], items[item_id]["name"], data["qty"], data["sales"].quantize(Decimal("0.01")), data["cost"].quantize(Decimal("0.01")), margin.quantize(Decimal("0.01")),
                         f"{(margin / data['sales'] * 100):.1f}%" if data["sales"] else ""])
        rows.append(["TOTAL", "", "", totals[0].quantize(Decimal("0.01")), totals[1].quantize(Decimal("0.01")), totals[2].quantize(Decimal("0.01")), f"{(totals[2] / totals[0] * 100):.1f}%" if totals[0] else ""])
        title = "Sales Margin (Cost of Goods Sold)"
        sections.append({"heading": f"Items issued {display_date(date_from)} to {display_date(date_to)} - sales at invoice price, cost at {'FIFO' if method == 'fifo' else 'weighted average'}",
                         "headers": ["Item Code", "Item", "Quantity Sold", f"Sales ({currency})", "Cost of Goods Sold", "Gross Margin", "Margin %"], "rows": rows, "total_rows": [len(rows) - 1]})
    elif report in ("reorder", "slow"):
        state = run_costing(database, date_to, method); rows = []
        cutoff = (datetime.strptime(date_to, "%Y-%m-%d") - timedelta(days=int(options.get("days") or 90))).strftime("%Y-%m-%d")
        for item_id, item in sorted(items.items(), key=lambda pair: pair[1]["sku"]):
            if not item["active"] or not in_category(item_id): continue
            data = state.get(item_id, {"qty": ZERO, "avg": ZERO, "last_out": None, "value": ZERO}); reorder = _d(item.get("reorder_level"))
            if report == "reorder" and reorder > 0 and data["qty"] <= reorder:
                rows.append([item["sku"], item["name"], item["unit"], data["qty"], reorder, max(ZERO, reorder * 2 - data["qty"]), data["avg"].quantize(Decimal("0.0001"))])
            if report == "slow" and data["qty"] > 0 and (not data.get("last_out") or data["last_out"] < cutoff):
                rows.append([item["sku"], item["name"], item["unit"], data["qty"], data["value"].quantize(Decimal("0.01")), display_date(data.get("last_out")) if data.get("last_out") else "never sold"])
        if report == "reorder":
            title = "Reorder Report"; sections.append({"heading": f"Items at or below their reorder level on {display_date(date_to)}", "headers": ["Item Code", "Item", "Unit", "On Hand", "Reorder Level", "Suggested Order", "Unit Cost"],
                                                          "rows": rows or [["No item below its reorder level"] + [""] * 6], "total_rows": []})
        else:
            title = "Slow-moving Stock"; sections.append({"heading": f"Items in stock with no issue since {display_date(cutoff)} ({int(options.get('days') or 90)} days)",
                                                          "headers": ["Item Code", "Item", "Unit", "On Hand", f"Value ({currency})", "Last Issue"], "rows": rows or [["No slow-moving items"] + [""] * 5], "total_rows": []})
    else: raise ValueError("Unknown inventory report")
    meta = [f"Company: {company.get('company_name') or '-'}   Inventory currency: {currency}   Costing: {'FIFO' if method == 'fifo' else 'Weighted average'}",
            f"Period: {display_date(date_from)} to {display_date(date_to)}" + (("   Filters: " + ", ".join(filters)) if filters else "")]
    return {"title": title, "meta": meta, "sections": sections}


# ---------------------------------------------------------------- year end
def stock_value(database, date_to, method=None):
    state = run_costing(database, iso_date(date_to), method)
    return sum((data["value"] for data in state.values()), ZERO).quantize(Decimal("0.01"))


def post_stock_variation(database, year, user_id):
    """Periodic method: cancel the opening stock (Dr 6051 / Cr 37) and book the closing stock (Dr 37 / Cr 6052)."""
    year = int(year); inv = settings(database); currency = inv["currency"]
    with database.connect() as db:
        for row in db.execute("SELECT id FROM journal_entries WHERE source_type='journal_voucher' AND voucher_type='06' AND description LIKE ?", (f"STOCK VARIATION - {year}%",)).fetchall():
            db.execute("DELETE FROM journal_entries WHERE id=?", (row["id"],))
    # The stock already in the ledger (account 37 at year end, before this voucher) is cancelled; the counted closing stock replaces it.
    opening = _ledger_stock(database, f"{year}-12-31", currency); closing = stock_value(database, f"{year}-12-31")
    if not opening and not closing: return {"year": year, "opening": 0.0, "closing": 0.0, "voucher": None}
    lines = []
    if opening > 0: lines += [{"account_code": OPENING_ACCOUNT, "line_currency": currency, "side": "D", "amount": str(opening)}, {"account_code": STOCK_ACCOUNT, "line_currency": currency, "side": "C", "amount": str(opening)}]
    elif opening < 0: lines += [{"account_code": STOCK_ACCOUNT, "line_currency": currency, "side": "D", "amount": str(-opening)}, {"account_code": OPENING_ACCOUNT, "line_currency": currency, "side": "C", "amount": str(-opening)}]
    if closing: lines += [{"account_code": STOCK_ACCOUNT, "line_currency": currency, "side": "D", "amount": str(closing)}, {"account_code": CLOSING_ACCOUNT, "line_currency": currency, "side": "C", "amount": str(closing)}]
    voucher = database.save_journal_voucher({"entry_date": f"31-12-{year}", "description": f"STOCK VARIATION - {year}: opening {opening:,.2f} / closing {closing:,.2f} {currency}",
                                             "currency": currency, "voucher_type": "06"}, lines, user_id)
    return {"year": year, "opening": float(opening), "closing": float(closing), "variation": float(closing - opening), "voucher": voucher["voucher"]["entry_number"]}


def _ledger_stock(database, date_to, currency):
    from ledger_reports import _load_lines, _digits
    column = currency if currency in ("LBP", "USD") else "account"
    total = sum((row["signed"][column] for row in _load_lines(database, {"posting_status": "posted"})
                 if _digits(row["code"]).startswith(STOCK_ACCOUNT) and row["iso_date"] <= date_to and (column != "account" or row["account_currency"] == currency)), ZERO)
    return total.quantize(Decimal("0.01"))


def _opening_value(database, year):
    value = stock_value(database, f"{int(year) - 1}-12-31")
    with database.connect() as db:
        rows = db.execute("""SELECT m.quantity,m.unit_cost FROM stock_movements m JOIN stock_documents d ON d.id=m.document_id
            WHERE d.doc_type='opening' AND d.doc_date>=? AND d.doc_date<=?""", (f"{int(year)}-01-01", f"{int(year)}-12-31")).fetchall()
    return (value + sum((_d(r["quantity"]) * _d(r["unit_cost"]) for r in rows), ZERO)).quantize(Decimal("0.01"))


def carry_forward(source, target, year, user_id):
    """New fiscal-year file: copy items and warehouses, and post the closing stock of year-1 as its Opening Stock."""
    year = int(year); state = run_costing(source, f"{year - 1}-12-31")
    with source.connect() as src: items = [dict(r) for r in src.execute("SELECT * FROM inventory_items")]; warehouses = [dict(r) for r in src.execute("SELECT * FROM warehouses")]
    with target.connect() as db:
        for w in warehouses: db.execute("INSERT OR IGNORE INTO warehouses(id,code,name,active) VALUES(?,?,?,?)", (w["id"], w["code"], w["name"], w["active"]))
        for i in items:
            db.execute("""INSERT INTO inventory_items(id,sku,name,unit,quantity,average_cost,category,reorder_level,sales_price,active,notes,barcode,created_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)
                ON CONFLICT(id) DO UPDATE SET sku=excluded.sku,name=excluded.name,unit=excluded.unit,category=excluded.category,reorder_level=excluded.reorder_level,sales_price=excluded.sales_price,active=excluded.active""",
                (i["id"], i["sku"], i["name"], i["unit"], "0", "0", i.get("category"), i.get("reorder_level") or "0", i.get("sales_price") or "0", i.get("active", 1), i.get("notes"), i.get("barcode"), i.get("created_at")))
        for old in db.execute("SELECT id FROM stock_documents WHERE doc_type='opening' AND number LIKE ?", (f"OPN-{year}-%",)).fetchall():
            db.execute("DELETE FROM stock_movements WHERE document_id=?", (old["id"],)); db.execute("DELETE FROM stock_documents WHERE id=?", (old["id"],))
    created = []
    by_warehouse = {}
    for item_id, data in state.items():
        for warehouse_id, qty in data["by_warehouse"].items():
            if qty > 0: by_warehouse.setdefault(warehouse_id, []).append({"item_id": item_id, "quantity": qty, "unit_cost": data["avg"].quantize(Decimal("0.0001"))})
    for warehouse_id, lines in by_warehouse.items():
        doc = save_document(target, {"doc_type": "opening", "doc_date": f"01-01-{year}", "warehouse_id": warehouse_id, "notes": f"Closing stock of {year - 1}"}, lines, user_id)
        created.append(doc["number"])
    return created


# ---------------------------------------------------------------- categories, units
def list_categories(database):
    with database.connect() as db:
        rows = [dict(r) for r in db.execute("SELECT * FROM item_categories ORDER BY name")]
        units = [r["name"] for r in db.execute("SELECT name FROM item_units ORDER BY name")]
        for name in {r["category"] for r in db.execute("SELECT DISTINCT category FROM inventory_items WHERE category IS NOT NULL AND category<>''")}:
            if not any(c["name"] == name and not c["parent_id"] for c in rows): rows.append({"id": None, "name": name, "parent_id": None})
    top = [c for c in rows if not c["parent_id"]]
    return {"categories": [{"id": c["id"], "name": c["name"], "subcategories": [s["name"] for s in rows if s["parent_id"] and s["parent_id"] == c["id"]]} for c in sorted(top, key=lambda c: c["name"])],
            "units": units}


def save_category(database, item, user_id):
    name = str(item.get("name") or "").strip(); kind = str(item.get("kind") or "category")
    if not name: raise ValueError("Enter the name")
    with database.connect() as db:
        if kind == "unit": db.execute("INSERT OR IGNORE INTO item_units(name) VALUES(?)", (name,))
        elif kind == "subcategory":
            parent = str(item.get("parent") or "").strip()
            if not parent: raise ValueError("Choose the category of the subcategory")
            db.execute("INSERT OR IGNORE INTO item_categories(name,parent_id) VALUES(?,NULL)", (parent,))
            parent_id = db.execute("SELECT id FROM item_categories WHERE name=? AND parent_id IS NULL", (parent,)).fetchone()["id"]
            db.execute("INSERT OR IGNORE INTO item_categories(name,parent_id) VALUES(?,?)", (name, parent_id))
        else:
            if not db.execute("SELECT 1 FROM item_categories WHERE name=? AND parent_id IS NULL", (name,)).fetchone(): db.execute("INSERT INTO item_categories(name,parent_id) VALUES(?,NULL)", (name,))
    return list_categories(database)


def find_or_create_item(database, name, unit="unit", code=None, user_id=None, supplier_id=None):
    """Used by the purchase import: an item that does not exist yet is created automatically."""
    with database.connect() as db:
        row = db.execute("SELECT * FROM inventory_items WHERE (sku=? AND ?<>'') OR lower(name)=lower(?) ORDER BY id LIMIT 1", (str(code or "").upper(), str(code or ""), str(name or ""))).fetchone()
    if row: return dict(row)
    return save_item(database, {"sku": code or "", "name": name, "unit": unit or "unit", "supplier_id": supplier_id}, user_id)


# ---------------------------------------------------------------- physical inventory
def count_sheet(database, warehouse_id, date):
    """System quantity of every active item in a warehouse on a date, ready for counting."""
    date = iso_date(date); state = run_costing(database, date); warehouse = int(warehouse_id)
    return [{"item_id": i["id"], "sku": i["sku"], "name": i["name"], "unit": i["unit"], "category": i.get("category") or "", "location": i.get("location") or "",
             "system_qty": float(state.get(i["id"], {}).get("by_warehouse", {}).get(warehouse, ZERO)), "unit_cost": float(state.get(i["id"], {}).get("avg", ZERO))}
            for i in list_items(database, include_inactive=False)]


def save_count(database, header, lines, user_id, count_id=None, post=False):
    date = iso_date(header.get("count_date"), "Count date"); warehouse = int(header.get("warehouse_id") or 0)
    if not warehouse: raise ValueError("Choose the warehouse")
    clean = []
    for line in lines or []:
        if line.get("counted") in (None, ""): continue
        counted = _d(line["counted"])
        if counted < 0: raise ValueError(f"{line.get('sku')}: the counted quantity cannot be negative")
        clean.append({"item_id": int(line["item_id"]), "sku": line.get("sku"), "counted": str(counted)})
    with database.connect() as db:
        if count_id:
            row = db.execute("SELECT * FROM physical_counts WHERE id=?", (int(count_id),)).fetchone()
            if not row: raise KeyError("Count not found")
            if row["status"] == "posted": raise ValueError("This count is already posted to the stock")
            db.execute("UPDATE physical_counts SET count_date=?,warehouse_id=?,lines=?,notes=? WHERE id=?", (date, warehouse, json.dumps(clean), header.get("notes"), int(count_id))); saved = int(count_id)
        else:
            numbers = [int(r["number"].rsplit("-", 1)[-1]) for r in db.execute("SELECT number FROM physical_counts WHERE number LIKE ?", (f"PHC-{date[:4]}-%",)) if r["number"].rsplit("-", 1)[-1].isdigit()]
            saved = db.execute("INSERT INTO physical_counts(number,count_date,warehouse_id,lines,notes,created_by,created_at) VALUES(?,?,?,?,?,?,?)",
                (f"PHC-{date[:4]}-{max(numbers, default=0) + 1:06d}", date, warehouse, json.dumps(clean), header.get("notes"), user_id, utcnow())).lastrowid
    if post:
        system = {l["item_id"]: l for l in count_sheet(database, warehouse, date)}; gains = []; losses = []
        for line in clean:
            difference = _d(line["counted"]) - _d(system.get(line["item_id"], {}).get("system_qty", 0))
            if difference > 0: gains.append({"item_id": line["item_id"], "quantity": difference, "unit_cost": system.get(line["item_id"], {}).get("unit_cost", 0) or 0})
            elif difference < 0: losses.append({"item_id": line["item_id"], "quantity": -difference})
        numbers = []
        with database.connect() as db: number = db.execute("SELECT number FROM physical_counts WHERE id=?", (saved,)).fetchone()["number"]
        if gains: numbers.append(save_document(database, {"doc_type": "adjustment_in", "doc_date": date, "warehouse_id": warehouse, "reference": number, "notes": f"Physical count {number}"}, gains, user_id)["number"])
        if losses: numbers.append(save_document(database, {"doc_type": "adjustment_out", "doc_date": date, "warehouse_id": warehouse, "reference": number, "notes": f"Physical count {number}"}, losses, user_id)["number"])
        with database.connect() as db: db.execute("UPDATE physical_counts SET status='posted',adjustment_numbers=? WHERE id=?", (", ".join(numbers), saved))
    return get_count(database, saved)


def get_count(database, count_id):
    with database.connect() as db:
        row = db.execute("SELECT c.*,w.code warehouse_code FROM physical_counts c JOIN warehouses w ON w.id=c.warehouse_id WHERE c.id=?", (int(count_id),)).fetchone()
    if not row: raise KeyError("Count not found")
    result = dict(row); result["lines"] = json.loads(result["lines"] or "[]"); return result


def list_counts(database):
    with database.connect() as db:
        return [dict(r) for r in db.execute("SELECT c.id,c.number,c.count_date,c.status,c.adjustment_numbers,w.code warehouse_code FROM physical_counts c JOIN warehouses w ON w.id=c.warehouse_id ORDER BY c.id DESC")]
