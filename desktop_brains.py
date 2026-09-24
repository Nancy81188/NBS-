"""BRAINS-style screens (version 1.13): Journal Voucher with multi-currency lines and the
Balance des Comptes panel used by both the Trial Balance and the Statement of Account."""
from __future__ import annotations

import os
import tempfile
import tkinter as tk
from datetime import datetime
from tkinter import messagebox, ttk

from report_export import export_sections_pdf

NAVY, GOLD, LIGHT = "#071b2e", "#c9a96a", "#f3f6f8"
RED, MUTED = "#8B1E1E", "#5f6b76"
VOUCHER_TYPES = ["01 - General Voucher", "02 - Receipt Voucher", "03 - Payment Voucher", "04 - Opening Voucher", "05 - Closing Voucher", "06 - Adjustment"]


def _num(value):
    try: return float(str(value or 0).replace(",", ""))
    except ValueError: return 0.0


def _fmt(value, places=2):
    return f"{_num(value):,.{places}f}"


def _date_text(value):
    text = str(value or "").strip()
    for pattern in ("%d-%m-%Y", "%d%m%Y", "%Y-%m-%d"):
        try: return datetime.strptime(text, pattern).strftime("%d-%m-%Y")
        except ValueError: pass
    return text


class EditableSheet:
    """A Treeview that edits like a spreadsheet: double-click / Enter to type, Tab / Enter to move on."""

    def __init__(self, app, parent, columns, editable, on_change, on_select=None, height=10, lookup_column=None):
        self.app = app; self.columns = columns; self.editable = editable; self.on_change = on_change; self.on_select = on_select
        self.lookup_column = lookup_column; self.rows = {}
        frame = tk.Frame(parent, bg=LIGHT); frame.pack(fill="both", expand=True, padx=10, pady=4)
        self.tree = ttk.Treeview(frame, columns=[c[0] for c in columns], show="headings", height=height, selectmode="browse")
        for key, label, width, anchor in columns: self.tree.heading(key, text=label); self.tree.column(key, width=width, anchor=anchor, stretch=key == "account")
        yscroll = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview); xscroll = ttk.Scrollbar(frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew"); yscroll.grid(row=0, column=1, sticky="ns"); xscroll.grid(row=1, column=0, sticky="ew")
        frame.grid_rowconfigure(0, weight=1); frame.grid_columnconfigure(0, weight=1)
        self.tree.tag_configure("odd", background="#fbf3e4")
        self.tree.bind("<Double-1>", self._clicked); self.tree.bind("<Return>", lambda _e: self.edit(self.tree.focus(), self.editable[0]))
        self.tree.bind("<<TreeviewSelect>>", lambda _e: self.on_select(self.selected()) if self.on_select else None)

    def clear(self):
        self.tree.delete(*self.tree.get_children()); self.rows = {}

    def insert(self, row, index="end"):
        iid = self.tree.insert("", index, values=self.values(row)); self.rows[iid] = row; self.renumber(); return iid

    def values(self, row):
        return [row.get("_display", {}).get(key, row.get(key, "")) for key, _l, _w, _a in self.columns]

    def refresh(self, iid):
        if self.tree.exists(iid): self.tree.item(iid, values=self.values(self.rows[iid]))

    def renumber(self):
        for number, iid in enumerate(self.tree.get_children(), 1):
            self.rows[iid]["line"] = f"{number:03d}"; self.refresh(iid)
            self.tree.item(iid, tags=("odd",) if number % 2 else ())

    def ordered(self):
        return [self.rows[iid] for iid in self.tree.get_children()]

    def selected(self):
        iid = self.tree.focus(); return (iid, self.rows.get(iid)) if iid else (None, None)

    def delete_selected(self):
        iid = self.tree.focus()
        if not iid: return False
        self.tree.delete(iid); self.rows.pop(iid, None); self.renumber(); return True

    def _clicked(self, event):
        iid = self.tree.identify_row(event.y); column = self.tree.identify_column(event.x)
        if not iid or not column: return
        key = self.columns[int(column.lstrip("#")) - 1][0]
        self.edit(iid, key if key in self.editable else self.editable[0])

    def edit(self, iid, key):
        if not iid or not self.tree.exists(iid) or not self.tree.winfo_ismapped(): return
        index = [c[0] for c in self.columns].index(key); self.tree.see(iid)
        bbox = self.tree.bbox(iid, f"#{index + 1}")
        if not bbox: return
        row = self.rows[iid]; value = row.get(key, "")
        editor = tk.Entry(self.tree, justify={"w": "left", "e": "right"}.get(self.columns[index][3], "center"))
        editor.insert(0, "" if value is None else str(value)); editor.place(x=bbox[0], y=bbox[1], width=max(bbox[2], 70), height=bbox[3])
        editor.focus_set(); editor.select_range(0, "end"); done = {"flag": False}
        def commit(move):
            if done["flag"]: return
            done["flag"] = True; text = editor.get().strip(); editor.destroy()
            if self.on_change(iid, key, text) is False: return
            self.refresh(iid)
            if move:
                position = self.editable.index(key)
                if position + 1 < len(self.editable): self.app.after(10, lambda: self.edit(iid, self.editable[position + 1]))
                else:
                    rows = self.tree.get_children(); at = rows.index(iid)
                    if at + 1 < len(rows): self.app.after(10, lambda: self.edit(rows[at + 1], self.editable[0]))
        editor.bind("<Return>", lambda _e: commit(True)); editor.bind("<Tab>", lambda _e: (commit(True), "break")[1])
        editor.bind("<FocusOut>", lambda _e: commit(False)); editor.bind("<Escape>", lambda _e: (done.__setitem__("flag", True), editor.destroy()))
        if key == self.lookup_column:
            def lookup(_e=None):
                variable = tk.StringVar(value=editor.get()); done["flag"] = True; editor.destroy()
                def chosen(*_args):
                    if variable.get() and self.on_change(iid, key, variable.get().split(" - ", 1)[0].strip()) is not False: self.refresh(iid)
                variable.trace_add("write", chosen); self.app.open_account_lookup(variable); return "break"
            editor.bind("<F2>", lookup)


class BrainsScreensMixin:
    # ================================================================ Journal Voucher
    def build_manual(self):
        page = self.manual_tab; self.editing_voucher_id = None; self.voucher_rates = {}
        bar = tk.Frame(page, bg=NAVY); bar.pack(fill="x", padx=10, pady=(8, 0))
        tk.Label(bar, text="General Voucher", bg=NAVY, fg="white", font=("Segoe UI", 11, "bold")).pack(side="left", padx=10, pady=5)
        for text, step in (("|<", "first"), ("<", "previous"), (">", "next"), (">|", "last")):
            tk.Button(bar, text=text, command=lambda s=step: self.navigate_voucher(s), bg=GOLD, fg=NAVY, border=0, width=3, font=("Segoe UI", 9, "bold")).pack(side="left", padx=2, pady=4)
        tk.Button(bar, text="New", command=self.new_manual_voucher, bg="white", fg=NAVY, border=0, padx=12).pack(side="left", padx=(12, 2), pady=4)
        tk.Button(bar, text="Save", command=self.save_manual_invoice, bg=GOLD, fg=NAVY, border=0, padx=14, font=("Segoe UI", 9, "bold")).pack(side="left", padx=2, pady=4)
        tk.Button(bar, text="Delete", command=self.delete_current_voucher, bg=RED, fg="white", border=0, padx=12).pack(side="left", padx=2, pady=4)
        for text, fmt in (("Excel", "xlsx"), ("PDF", "pdf"), ("Print", "print")):
            tk.Button(bar, text=text, command=lambda f=fmt: self.manual_entry_report(f), bg="white", fg=NAVY, border=0, padx=10).pack(side="right", padx=2, pady=4)
        header = tk.Frame(page, bg=LIGHT); header.pack(fill="x", padx=10, pady=6)
        self.manual_type = tk.StringVar(value=VOUCHER_TYPES[0]); self.manual_no = tk.StringVar(); self.manual_date = tk.StringVar(value=datetime.now().strftime("%d-%m-%Y"))
        self.manual_currency = tk.StringVar(value="USD"); self.manual_find = tk.StringVar()
        tk.Label(header, text="Type", bg=LIGHT).pack(side="left"); ttk.Combobox(header, textvariable=self.manual_type, values=VOUCHER_TYPES, state="readonly", width=17).pack(side="left", padx=(4, 8))
        tk.Label(header, text="Number", bg=LIGHT).pack(side="left")
        tk.Entry(header, textvariable=self.manual_no, width=15, state="readonly", readonlybackground="white", font=("Segoe UI", 10, "bold")).pack(side="left", padx=(4, 8))
        tk.Label(header, text="Date", bg=LIGHT).pack(side="left"); self.date_entry(header, self.manual_date, 12).pack(side="left", padx=(4, 12))
        tk.Label(header, text="Currency", bg=LIGHT).pack(side="left")
        ttk.Combobox(header, textvariable=self.manual_currency, values=["USD", "LBP", "EUR", "AED"], state="readonly", width=6).pack(side="left", padx=(4, 12))
        tk.Label(header, text="Branch", bg=LIGHT).pack(side="left"); self.branch_selector(header, self.manual_branch, 12, False).pack(side="left", padx=(4, 8))
        tk.Label(header, text="Find", bg=LIGHT).pack(side="left")
        self.manual_find_box = ttk.Combobox(header, textvariable=self.manual_find, width=24); self.manual_find_box.pack(side="left", padx=4)
        self.manual_find_box.bind("<<ComboboxSelected>>", lambda _e: self.open_found_voucher()); self.manual_find_box.bind("<KeyRelease>", self.search_vouchers)
        self.manual_currency.trace_add("write", lambda *_a: self.update_manual_totals())
        self.manual_date.trace_add("write", lambda *_a: self.voucher_date_changed())
        columns = [("line", "#", 45, "center"), ("account", "Account No.", 110, "w"), ("line_currency", "Currency", 70, "center"), ("side", "D/C", 45, "center"),
                   ("amount", "Amount (Account Currency)", 165, "e"), ("amount_lbp", "Amount LBP", 145, "e"), ("amount_usd", "Amount USD", 120, "e"),
                   ("due_date", "Due Date", 95, "center"), ("reference", "Reference", 110, "w"), ("department", "Dep.", 60, "center"), ("project", "Project", 85, "center"),
                   ("rate_lbp", "Rate LBP", 95, "e"), ("rate_usd", "Rate USD", 95, "e")]
        bottom = tk.Frame(page, bg=LIGHT); bottom.pack(side="bottom", fill="x", padx=10, pady=(2, 6))
        totals = tk.LabelFrame(bottom, text="Totals", bg=LIGHT, padx=10, pady=4); totals.pack(side="right")
        tk.Label(totals, text="", bg=LIGHT).grid(row=0, column=0)
        for column, text in enumerate(("LBP", "USD", "Voucher Cur."), 1): tk.Label(totals, text=text, bg=LIGHT, font=("Segoe UI", 9, "bold")).grid(row=0, column=column, padx=6)
        self.voucher_total_labels = {}
        for row, name in enumerate(("Debit", "Credit", "Balance"), 1):
            tk.Label(totals, text=name, bg=LIGHT).grid(row=row, column=0, sticky="e", padx=4)
            for column, key in enumerate(("lbp", "usd", "voucher"), 1):
                label = tk.Label(totals, text="0.00", bg="#dfe6ee", width=15, anchor="e", font=("Segoe UI", 9, "bold")); label.grid(row=row, column=column, padx=3, pady=2)
                self.voucher_total_labels[(name, key)] = label
        left = tk.Frame(bottom, bg=LIGHT); left.pack(side="left", fill="both", expand=True)
        tk.Label(left, text="Details", bg=LIGHT, font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="nw", padx=(0, 6))
        self.manual_details = tk.Text(left, width=40, height=3, font=("Segoe UI", 9)); self.manual_details.grid(row=0, column=1, sticky="w")
        line_buttons = tk.Frame(left, bg=LIGHT); line_buttons.grid(row=1, column=1, sticky="w", pady=3)
        self.action_button(line_buttons, "Add Line", self.add_manual_item).pack(side="left", padx=(0, 4))
        tk.Button(line_buttons, text="Delete Line", command=self.remove_manual_item, bg=RED, fg="white", border=0, padx=12, pady=7).pack(side="left", padx=4)
        self.action_button(line_buttons, "Insert Line", self.insert_manual_item).pack(side="left", padx=4)
        tk.Label(left, text="Double-click a cell to type · F2 = account list · D/C: D or C · Currency: USD, LBP, EUR, AED", bg=LIGHT, fg=MUTED, wraplength=420, justify="left").grid(row=2, column=0, columnspan=2, sticky="w")
        self.manual_line_info = tk.Label(page, text="", bg="#dfe6ee", fg=NAVY, anchor="w", font=("Segoe UI", 9, "bold"), padx=8)
        self.manual_line_info.pack(side="bottom", fill="x", padx=10)
        self.voucher_sheet = EditableSheet(self, page, columns, ["account", "line_currency", "side", "amount", "due_date", "reference", "department", "project", "rate_lbp", "rate_usd"],
                                           self.voucher_cell_changed, self.voucher_line_selected, height=6, lookup_column="account")
        self.manual_items = []; self.manual_tree = self.voucher_sheet.tree
        self.load_manual_vouchers(); self.new_manual_voucher(confirm=False)

    # ---- rates and lines
    def voucher_rates_for(self, currency):
        key = (currency, self.manual_date.get().strip())
        if key not in self.voucher_rates:
            try: self.voucher_rates[key] = self.client.suggested_rates(currency, key[1] if len(key[1]) == 10 else None)
            except Exception: self.voucher_rates[key] = {"rate_lbp": 1 if currency == "LBP" else 89500, "rate_usd": 89500 if currency == "LBP" else 1}
        return self.voucher_rates[key]

    def voucher_date_changed(self):
        if len(self.manual_date.get()) == 10 and not self.editing_voucher_id: self.set_next_manual_voucher_number()

    def new_voucher_line(self, account=""):
        currency = self.manual_currency.get() or "USD"; rates = self.voucher_rates_for(currency)
        return self.recalculate_voucher_line({"account": account, "line_currency": currency, "side": "D", "amount": "", "due_date": self.manual_date.get(), "reference": "", "department": "", "project": "",
                                              "rate_lbp": rates["rate_lbp"], "rate_usd": rates["rate_usd"]})

    def recalculate_voucher_line(self, row):
        amount = _num(row.get("amount")); rate_lbp = _num(row.get("rate_lbp")) or 1; rate_usd = _num(row.get("rate_usd")) or 1
        row["amount_lbp"] = amount * rate_lbp; row["amount_usd"] = amount / rate_usd if row.get("line_currency") == "LBP" else amount * rate_usd
        row["_display"] = {"amount": _fmt(amount, 3) if row.get("amount") not in ("", None) else "", "amount_lbp": _fmt(row["amount_lbp"], 3), "amount_usd": _fmt(row["amount_usd"], 3),
                           "rate_lbp": f"{rate_lbp:,.4f}", "rate_usd": f"{rate_usd:,.4f}"}
        return row

    def voucher_cell_changed(self, iid, key, text):
        row = self.voucher_sheet.rows[iid]
        if key == "account":
            code = text.split(" - ", 1)[0].strip()
            if code:
                account = self.account_by_code(code)
                if not account: messagebox.showwarning("Journal Voucher", f"Account {code} was not found. Press F2 in the cell to search."); return False
                row["account_name"] = account["name_en"]
                party = next((p for p in getattr(self, "party_rows", []) or [] if p.get("account_number") == code), None)
                if party and party.get("currency") and not row.get("amount"): self.voucher_cell_changed(iid, "line_currency", party["currency"])
            row["account"] = code
        elif key == "line_currency":
            currency = text.upper() or self.manual_currency.get()
            if currency in ("01", "1"): currency = "LBP"
            if currency in ("02", "2"): currency = "USD"
            if currency not in ("USD", "LBP", "EUR", "AED"): messagebox.showwarning("Journal Voucher", "Currency must be USD, LBP, EUR or AED"); return False
            rates = self.voucher_rates_for(currency); row.update(line_currency=currency, rate_lbp=rates["rate_lbp"], rate_usd=rates["rate_usd"])
        elif key == "side":
            side = text.upper()[:1]
            if side not in ("D", "C"): messagebox.showwarning("Journal Voucher", "Type D for Debit or C for Credit"); return False
            row["side"] = side
        elif key in ("amount", "rate_lbp", "rate_usd"):
            if text and _num(text) <= 0 and text.replace(",", "") not in ("0",):
                messagebox.showwarning("Journal Voucher", "Enter a positive number"); return False
            row[key] = text.replace(",", "")
        elif key == "due_date": row["due_date"] = _date_text(text)
        elif key in ("department", "project"):
            code = text.split(" - ", 1)[0].strip().upper()
            items = self.dimension_lists(refresh=True)["departments" if key == "department" else "projects"]
            if code and not any(i["code"].upper() == code for i in items):
                messagebox.showwarning("Journal Voucher", f"{key.title()} {code} was not found. Available: " + ", ".join(i["code"] for i in items[:15])); return False
            row[key] = code
        else: row[key] = text
        self.recalculate_voucher_line(row); self.update_manual_totals(); self.voucher_line_selected((iid, row))
        rows = self.voucher_sheet.tree.get_children()
        if key == "reference" and rows and iid == rows[-1] and row.get("account") and row.get("amount"): self.add_manual_item(edit=False)

    def account_by_code(self, code):
        if not getattr(self, "_account_cache", None):
            try: self._account_cache = {str(a["code"]): a for a in self.client.accounts()}
            except Exception: self._account_cache = {}
        return self._account_cache.get(str(code))

    def voucher_line_selected(self, selection):
        _iid, row = selection
        if not row: self.manual_line_info.config(text=""); return
        name = row.get("account_name") or (self.account_by_code(row.get("account")) or {}).get("name_en", "")
        self.manual_line_info.config(text=f"{row.get('account') or ''}  {name}      {_fmt(row.get('amount'), 3)} {row.get('line_currency')}      {_fmt(row.get('amount_lbp'), 3)} LBP      {_fmt(row.get('amount_usd'), 3)} USD")

    def add_manual_item(self, edit=True):
        iid = self.voucher_sheet.insert(self.new_voucher_line())
        self.voucher_sheet.tree.selection_set(iid); self.voucher_sheet.tree.focus(iid)
        if edit: self.after(30, lambda: self.voucher_sheet.edit(iid, "account"))
        self.update_manual_totals()

    def insert_manual_item(self):
        iid, _row = self.voucher_sheet.selected()
        index = self.voucher_sheet.tree.index(iid) if iid else "end"
        new = self.voucher_sheet.insert(self.new_voucher_line(), index)
        self.voucher_sheet.tree.selection_set(new); self.voucher_sheet.tree.focus(new); self.after(30, lambda: self.voucher_sheet.edit(new, "account"))

    def remove_manual_item(self):
        if not self.voucher_sheet.delete_selected(): return messagebox.showwarning("Journal Voucher", "Select a line first")
        self.update_manual_totals()

    def voucher_lines(self):
        return [row for row in self.voucher_sheet.ordered() if row.get("account") and _num(row.get("amount")) > 0]

    def voucher_value(self, row):
        currency = self.manual_currency.get()
        if row.get("line_currency") == currency: return _num(row.get("amount"))
        return row.get("amount_usd", 0) if currency == "USD" else row.get("amount_lbp", 0) if currency == "LBP" else None

    def update_manual_totals(self):
        if not hasattr(self, "voucher_total_labels"): return 0, 0
        totals = {key: {"D": 0.0, "C": 0.0} for key in ("lbp", "usd", "voucher")}; mixed = False
        for row in self.voucher_lines():
            side = row.get("side", "D"); totals["lbp"][side] += row.get("amount_lbp", 0); totals["usd"][side] += row.get("amount_usd", 0)
            value = self.voucher_value(row)
            if value is None: mixed = True
            else: totals["voucher"][side] += value
        for key in totals:
            debit, credit = totals[key]["D"], totals[key]["C"]; balance = debit - credit
            self.voucher_total_labels[("Debit", key)].config(text=_fmt(debit)); self.voucher_total_labels[("Credit", key)].config(text=_fmt(credit))
            self.voucher_total_labels[("Balance", key)].config(text="MIXED" if key == "voucher" and mixed else _fmt(balance), fg=NAVY if abs(balance) < 0.005 and not (key == "voucher" and mixed) else RED)
        self.manual_items = self.voucher_lines()
        return totals["voucher"]["D"], totals["voucher"]["C"]

    # ---- voucher list, navigation, open, save
    def load_manual_vouchers(self):
        if not hasattr(self, "manual_find_box"): return
        try: rows = [row for row in self.client.journal() if row.get("source_type") == "journal_voucher"]
        except Exception: rows = []
        grouped = {}
        for row in rows:
            item = grouped.setdefault(row["entry_id"], {"id": row["entry_id"], "number": row["entry_number"], "date": row["entry_date"], "description": row.get("description") or "",
                                                        "currency": row["currency"], "debit": 0.0})
            item["debit"] += float(row["debit"] or 0)
        self.manual_voucher_rows = grouped
        def key(item):
            try: return (datetime.strptime(_date_text(item["date"]), "%d-%m-%Y"), item["number"])
            except ValueError: return (datetime.min, item["number"])
        self.voucher_order = [item["id"] for item in sorted(grouped.values(), key=key)]
        self.voucher_choices = {f'{v["number"]} | {_date_text(v["date"])} | {v["description"][:40]} | {v["debit"]:,.2f} {v["currency"]}': v["id"] for v in grouped.values()}
        self.manual_find_box["values"] = list(self.voucher_choices)
        self.set_next_manual_voucher_number()

    def populate_manual_vouchers(self): self.load_manual_vouchers()

    def search_vouchers(self, _event=None):
        from desktop import row_matches_search
        typed = self.manual_find.get().strip(); choices = list(getattr(self, "voucher_choices", {}))
        self.manual_find_box["values"] = [c for c in choices if row_matches_search((c,), typed)] if typed else choices

    def set_next_manual_voucher_number(self):
        if not hasattr(self, "manual_no") or self.editing_voucher_id: return
        try: year = datetime.strptime(self.manual_date.get(), "%d-%m-%Y").year
        except ValueError: year = datetime.now().year
        prefix = f"JV-{year}-"; numbers = []
        for row in getattr(self, "manual_voucher_rows", {}).values():
            if str(row["number"]).startswith(prefix):
                try: numbers.append(int(str(row["number"]).rsplit("-", 1)[-1]))
                except ValueError: pass
        self.manual_no.set(f"{prefix}{max(numbers, default=0) + 1:06d}")

    def new_manual_voucher(self, confirm=True):
        if confirm and self.voucher_lines() and not self.editing_voucher_id and not messagebox.askyesno("Journal Voucher", "Start a new voucher? Lines that are not saved will be cleared."): return
        self.editing_voucher_id = None; self.voucher_sheet.clear(); self.manual_details.delete("1.0", "end"); self.manual_find.set("")
        self.manual_type.set(VOUCHER_TYPES[0]); self.manual_currency.set("USD"); self.manual_date.set(datetime.now().strftime("%d-%m-%Y"))
        self._account_cache = None; self.set_next_manual_voucher_number()
        for _ in range(2): self.voucher_sheet.insert(self.new_voucher_line())
        self.update_manual_totals(); self.manual_line_info.config(text="New voucher")

    def open_found_voucher(self):
        entry_id = getattr(self, "voucher_choices", {}).get(self.manual_find.get())
        if entry_id: self.open_voucher(entry_id)

    def navigate_voucher(self, step):
        order = getattr(self, "voucher_order", [])
        if not order: return messagebox.showinfo("Journal Voucher", "There are no saved vouchers yet")
        if step == "first": target = order[0]
        elif step == "last": target = order[-1]
        else:
            current = order.index(self.editing_voucher_id) if self.editing_voucher_id in order else (len(order) if step == "previous" else -1)
            position = current - 1 if step == "previous" else current + 1
            if not 0 <= position < len(order): return
            target = order[position]
        self.open_voucher(target)

    def open_voucher(self, entry_id):
        try: detail = self.client.journal_voucher(int(entry_id))
        except Exception as exc: return messagebox.showerror("Journal Voucher", str(exc))
        voucher = detail["voucher"]; self.editing_voucher_id = int(voucher["id"])
        self.manual_no.set(voucher["entry_number"]); self.manual_date.set(_date_text(voucher["entry_date"])); self.manual_currency.set(voucher["currency"])
        self.manual_type.set(next((t for t in VOUCHER_TYPES if t.startswith(str(voucher.get("voucher_type") or "01"))), VOUCHER_TYPES[0]))
        self.manual_details.delete("1.0", "end"); self.manual_details.insert("1.0", voucher.get("description") or "")
        try: self.manual_branch.set(next(b["name"] for b in self.client.branches() if b["id"] == voucher.get("branch_id")))
        except Exception: self.manual_branch.set("Head Office")
        self.voucher_sheet.clear()
        for line in detail["lines"]:
            side = "D" if float(line.get("debit") or 0) else "C"
            if line.get("line_currency"):
                row = {"account": line["account_code"], "account_name": line["account_name"], "line_currency": line["line_currency"], "side": side, "amount": line["amount"],
                       "rate_lbp": line["rate_lbp"], "rate_usd": line["rate_usd"], "due_date": line.get("due_date") or "", "reference": line.get("reference") or "",
                       "department": line.get("department") or "", "project": line.get("project") or ""}
            else:
                rates = self.voucher_rates_for(voucher["currency"])
                row = {"account": line["account_code"], "account_name": line["account_name"], "line_currency": voucher["currency"], "side": side,
                       "amount": float(line.get("debit") or 0) or float(line.get("credit") or 0), "rate_lbp": rates["rate_lbp"], "rate_usd": rates["rate_usd"], "due_date": "", "reference": "",
                       "department": line.get("department") or "", "project": line.get("project") or ""}
            self.voucher_sheet.insert(self.recalculate_voucher_line(row))
        self.update_manual_totals(); self.manual_line_info.config(text=f"Voucher {voucher['entry_number']} opened")

    def save_manual_invoice(self):
        lines = self.voucher_lines(); debit, credit = self.update_manual_totals()
        if len(lines) < 2: return messagebox.showwarning("Journal Voucher", "Enter at least two lines with an account and an amount")
        if any(self.voucher_value(row) is None for row in lines):
            return messagebox.showerror("Journal Voucher", f"A {self.manual_currency.get()} voucher can only contain {self.manual_currency.get()} lines. Choose USD or LBP as the voucher currency to mix currencies.")
        if abs(debit - credit) >= 0.005:
            needed = f"Credit {debit - credit:,.2f}" if debit > credit else f"Debit {credit - debit:,.2f}"
            return messagebox.showerror("Unbalanced Journal Voucher", f"Debit: {debit:,.2f}\nCredit: {credit:,.2f}\nStill needed: {needed} {self.manual_currency.get()}\n\nDebit must equal Credit before saving.")
        details = self.manual_details.get("1.0", "end").strip()
        if not details: return messagebox.showwarning("Journal Voucher", "Write the voucher details (description)")
        try: entry_date = datetime.strptime(self.manual_date.get().strip(), "%d-%m-%Y").strftime("%d-%m-%Y")
        except ValueError: return messagebox.showwarning("Journal Voucher", "Enter the date as 8 digits: DDMMYYYY")
        voucher = {"entry_number": self.manual_no.get().strip(), "entry_date": entry_date, "description": details, "currency": self.manual_currency.get(),
                   "branch": self.manual_branch.get(), "voucher_type": self.manual_type.get()[:2]}
        payload = [{"account_code": r["account"], "line_currency": r["line_currency"], "side": r["side"], "amount": r["amount"], "rate_lbp": r["rate_lbp"], "rate_usd": r["rate_usd"],
                    "due_date": r.get("due_date") or "", "reference": r.get("reference") or "", "department": r.get("department") or "", "project": r.get("project") or "",
                    "description": details.splitlines()[0][:120]} for r in lines]
        try: saved = self.client.save_journal_voucher(voucher, payload, self.editing_voucher_id)
        except Exception as exc: return messagebox.showerror("Journal Voucher", str(exc))
        messagebox.showinfo("Journal Voucher", f'Voucher {saved["voucher"]["entry_number"]} saved')
        self.load_manual_vouchers(); self.open_voucher(saved["voucher"]["id"]); self.load_journal(); self.load_trial()

    def delete_current_voucher(self):
        if not self.editing_voucher_id: return messagebox.showwarning("Journal Voucher", "Open a saved voucher first")
        if not messagebox.askyesno("Delete Journal Voucher", f"Delete voucher {self.manual_no.get()} and all its lines?"): return
        try: self.client.delete_journal_voucher(self.editing_voucher_id)
        except Exception as exc: return messagebox.showerror("Journal Voucher", str(exc))
        self.editing_voucher_id = None; self.load_manual_vouchers(); self.new_manual_voucher(confirm=False); self.load_journal(); self.load_trial()

    def delete_selected_manual_from_tab(self): self.delete_current_voucher()
    def edit_selected_manual_voucher(self): self.open_found_voucher()

    def manual_entry_report(self, format_name):
        lines = self.voucher_lines()
        if not lines: return messagebox.showwarning("Journal Voucher", "No lines to export")
        headers = ["#", "Account", "Account Name", "Currency", "D/C", "Amount", "Amount LBP", "Amount USD", "Due Date", "Reference", "Rate LBP", "Rate USD"]
        rows = [[r["line"], r["account"], (self.account_by_code(r["account"]) or {}).get("name_en", ""), r["line_currency"], r["side"], _num(r["amount"]), round(r["amount_lbp"], 2),
                 round(r["amount_usd"], 3), r.get("due_date", ""), r.get("reference", ""), _num(r["rate_lbp"]), _num(r["rate_usd"])] for r in lines]
        debit_lbp = sum(r["amount_lbp"] for r in lines if r["side"] == "D"); credit_lbp = sum(r["amount_lbp"] for r in lines if r["side"] == "C")
        debit_usd = sum(r["amount_usd"] for r in lines if r["side"] == "D"); credit_usd = sum(r["amount_usd"] for r in lines if r["side"] == "C")
        totals = [["", "", "Debit", "", "", "", round(debit_lbp, 2), round(debit_usd, 3), "", "", "", ""], ["", "", "Credit", "", "", "", round(credit_lbp, 2), round(credit_usd, 3), "", "", "", ""],
                  ["", "", "Balance", "", "", "", round(debit_lbp - credit_lbp, 2), round(debit_usd - credit_usd, 3), "", "", "", ""]]
        details = self.manual_details.get("1.0", "end").strip()
        title = f"Journal Voucher {self.manual_no.get()}"
        meta = [f"Type: {self.manual_type.get()}   Date: {self.manual_date.get()}   Voucher currency: {self.manual_currency.get()}", f"Details: {details}"]
        sections = [{"heading": "Voucher lines", "headers": headers, "rows": rows + totals, "total_rows": [len(rows), len(rows) + 1, len(rows) + 2]}]
        self.output_sections(title, meta, sections, title.replace(" ", "_"), format_name)

    def output_sections(self, title, meta, sections, name, format_name):
        if format_name == "print":
            handle = tempfile.NamedTemporaryFile(prefix="SaberAccounting_", suffix=".pdf", delete=False); handle.close()
            try:
                export_sections_pdf(handle.name, title, meta, sections)
                if os.name != "nt": raise RuntimeError("Printing is available in the Windows application")
                os.startfile(handle.name, "print")
            except Exception as exc: messagebox.showerror(title, str(exc))
            return
        self.save_sections(title, meta, sections, name, format_name)

    # ================================================================ Balance des Comptes
    def build_balance_panel(self, page, statement=False):
        year = getattr(self, "current_fiscal_year", datetime.now().year)
        v = {"account_from": tk.StringVar(), "account_to": tk.StringVar(), "date_from": tk.StringVar(value=f"01-01-{year}"), "date_to": tk.StringVar(value=f"31-12-{year}"),
             "print_date": tk.StringVar(value=datetime.now().strftime("%d-%m-%Y")), "branch": tk.StringVar(value="All Branches"), "summary_digits": tk.StringVar(value="4"),
             "posting": tk.StringVar(value="Posted only"), "first_column": tk.StringVar(value="account"), "second_column": tk.StringVar(value="LBP"), "party": tk.StringVar()}
        flags = {name: tk.BooleanVar(value=default) for name, default in (("summary", False), ("by_due_date", False), ("reference", statement), ("with_branch", False),
                 ("detailed", statement), ("include_zero", False), ("order_by_description", False), ("non_zero_only", False), ("chapters", False), ("sub_chapters", False),
                 ("balance_sheet_only", False), ("profit_loss_only", False), ("balance_format", False), ("carry_forward", True), ("monthly", False))}
        currencies = {code: tk.BooleanVar(value=True) for code in ("LBP", "USD", "EUR", "AED")}
        box = tk.LabelFrame(page, text="Statement of Account - options" if statement else "Balance des Comptes - options", bg=LIGHT, padx=8, pady=6)
        box.pack(fill="x", padx=10, pady=(8, 4))
        row0 = tk.Frame(box, bg=LIGHT); row0.pack(fill="x")
        if statement:
            tk.Label(row0, text="Customer / Supplier", bg=LIGHT, font=("Segoe UI", 9, "bold")).pack(side="left")
            party_box = ttk.Combobox(row0, textvariable=v["party"], width=34); party_box.pack(side="left", padx=(4, 12)); v["party_box"] = party_box
            party_box.bind("<<ComboboxSelected>>", lambda _e: self.balance_party_chosen(v)); party_box.bind("<KeyRelease>", lambda _e: self.balance_party_search(v))
        tk.Label(row0, text="Account From", bg=LIGHT).pack(side="left"); self.account_search_box(row0, v["account_from"], 14).pack(side="left", padx=(4, 8))
        tk.Label(row0, text="To", bg=LIGHT).pack(side="left"); self.account_search_box(row0, v["account_to"], 14).pack(side="left", padx=(4, 12))

        row1 = tk.Frame(box, bg=LIGHT); row1.pack(fill="x", pady=(4, 0))
        tk.Label(row1, text="Date From", bg=LIGHT).pack(side="left"); self.date_entry(row1, v["date_from"], 11).pack(side="left", padx=(4, 8))
        tk.Label(row1, text="To", bg=LIGHT).pack(side="left"); self.date_entry(row1, v["date_to"], 11).pack(side="left", padx=(4, 8))
        tk.Label(row1, text="Print Date", bg=LIGHT).pack(side="left"); self.date_entry(row1, v["print_date"], 11).pack(side="left", padx=(4, 8))
        tk.Label(row1, text="Branch", bg=LIGHT).pack(side="left"); self.branch_selector(row1, v["branch"], 14, True).pack(side="left", padx=(4, 8))
        ttk.Combobox(row1, textvariable=v["posting"], values=["Posted only", "Posted + Review", "Review only"], state="readonly", width=15).pack(side="left", padx=4)

        options = tk.Frame(box, bg=LIGHT); options.pack(fill="x", pady=(6, 0))
        groups = [("Lines", [("summary", "Summary (Resume)"), ("by_due_date", "By Due Date"), ("reference", "Reference"), ("with_branch", "With Branch")]),
                  ("Accounts", [("detailed", "Detailed Account (statement)"), ("include_zero", "All accounts"), ("order_by_description", "Order by Description"), ("non_zero_only", "Non-zero Balances only")]),
                  ("Grouping", [("chapters", "Chapters (class)"), ("sub_chapters", "Sub-chapters"), ("balance_sheet_only", "Balance Sheet (1-5)"), ("profit_loss_only", "Profit & Loss (6-7)")]),
                  ("Format", [("balance_format", "Format Balance (Dr / Cr balance)"), ("carry_forward", "With Carry Forward (opening)"), ("monthly", "Monthly")])]
        for title, items in groups:
            frame = tk.LabelFrame(options, text=title, bg=LIGHT, padx=4); frame.pack(side="left", fill="y", padx=(0, 6))
            for name, label in items: tk.Checkbutton(frame, text=label, variable=flags[name], bg=LIGHT, anchor="w").pack(anchor="w")
            if title == "Grouping":
                digits = tk.Frame(frame, bg=LIGHT); digits.pack(anchor="w")
                tk.Label(digits, text="Summary digits", bg=LIGHT).pack(side="left"); ttk.Combobox(digits, textvariable=v["summary_digits"], values=["1", "2", "3", "4", "5", "6"], width=3, state="readonly").pack(side="left", padx=3)
        row3 = tk.Frame(box, bg=LIGHT); row3.pack(fill="x", pady=(4, 0))
        for title, key, choices in (("1st Column", "first_column", (("account", "Account Currency"), ("LBP", "LBP"), ("USD", "USD"))),
                                    ("2nd Column", "second_column", (("account", "Account Currency"), ("LBP", "LBP"), ("USD", "USD"), ("none", "None")))):
            frame = tk.LabelFrame(row3, text=title, bg=LIGHT, padx=4); frame.pack(side="left", padx=(0, 6))
            for value, label in choices: tk.Radiobutton(frame, text=label, value=value, variable=v[key], bg=LIGHT).pack(side="left")
        actions = tk.Frame(row3, bg=LIGHT); actions.pack(side="right", padx=6)
        state = {"vars": v, "flags": flags, "currencies": currencies, "statement": statement, "result": None}
        self.add_dimension_options(state, box)
        tk.Label(state["dimension_row"], text="Currencies", bg=LIGHT).pack(side="left", padx=(10, 0))
        for code, var in currencies.items(): tk.Checkbutton(state["dimension_row"], text=code, variable=var, bg=LIGHT).pack(side="left")
        tk.Button(actions, text="Show", command=lambda: self.run_balance_report(state), bg=GOLD, fg=NAVY, border=0, padx=22, pady=7, font=("Segoe UI", 10, "bold")).pack(side="left", padx=3)
        for text, fmt in (("Print", "print"), ("Excel", "xlsx"), ("PDF", "pdf")):
            self.action_button(actions, text, lambda f=fmt: self.export_balance_report(state, f)).pack(side="left", padx=3)
        state["info"] = tk.Label(page, text="Choose the options and press Show.", bg=LIGHT, fg=MUTED, anchor="w"); state["info"].pack(fill="x", padx=12)
        state["viewer"] = self.report_viewer(page)
        return state

    def balance_party_search(self, v):
        from desktop import row_matches_search
        typed = v["party"].get().strip(); names = list(getattr(self, "balance_party_map", {}))
        v["party_box"]["values"] = [n for n in names if row_matches_search((n,), typed)] if typed else names

    def balance_party_chosen(self, v):
        party = getattr(self, "balance_party_map", {}).get(v["party"].get())
        if party and party.get("account_number"): v["account_from"].set(party["account_number"]); v["account_to"].set(party["account_number"])

    def balance_options(self, state):
        v, flags = state["vars"], state["flags"]
        options = {name: var.get() for name, var in flags.items()}
        if options.pop("include_zero"): options["non_zero_only"] = False
        options.update(account_from=v["account_from"].get().split(" - ", 1)[0].strip(), account_to=v["account_to"].get().split(" - ", 1)[0].strip(),
                       date_from=v["date_from"].get().strip(), date_to=v["date_to"].get().strip(), print_date=v["print_date"].get().strip(),
                       first_column=v["first_column"].get(), second_column=v["second_column"].get(), summary_digits=v["summary_digits"].get(),
                       currencies=[code for code, var in state["currencies"].items() if var.get()], statement=state["statement"],
                       posting_status={"Posted only": "posted", "Posted + Review": "all", "Review only": "review"}[v["posting"].get()],
                       department=self.dimension_code(v["department"].get()), project=self.dimension_code(v["project"].get()))
        if len(options["currencies"]) == 4: options["currencies"] = []
        branch = self.selected_branch_id(v["branch"])
        if branch: options["branch_id"] = branch
        for key in ("date_from", "date_to", "print_date"):
            if options[key]:
                try: datetime.strptime(options[key], "%d-%m-%Y")
                except ValueError: raise ValueError(f"{key.replace('_', ' ').title()} must be DD-MM-YYYY")
        return options

    def run_balance_report(self, state):
        try:
            if state["statement"] and not state["vars"]["account_from"].get().strip():
                raise ValueError("Choose a customer / supplier (or an account range) first")
            options = self.balance_options(state); result = self.client.account_report(options)
        except Exception as exc: return messagebox.showerror("Statement of Account" if state["statement"] else "Trial Balance", str(exc))
        state["result"] = result; self.show_sections(state["viewer"], result["sections"])
        state["info"].config(text=f"{result['title']}  |  {result['account_count']} account(s)  |  " + "   ".join(result["meta"][1:3]), fg=NAVY)

    def export_balance_report(self, state, format_name):
        if not state.get("result"): self.run_balance_report(state)
        result = state.get("result")
        if not result: return
        name = ("Statement_" + (state["vars"]["party"].get() or state["vars"]["account_from"].get())) if state["statement"] else "Trial_Balance"
        self.output_sections(result["title"], result["meta"], result["sections"], name.replace(" ", "_")[:60], format_name)

    # ---- tabs
    def build_trial(self):
        self.trial_state = self.build_balance_panel(self.trial_tab, statement=False)
        self.trial_rows = []

    def load_trial(self):
        state = getattr(self, "trial_state", None)
        if state and state.get("result"): self.run_balance_report(state)

    def build_statement(self):
        self.statement_state = self.build_balance_panel(self.statement_tab, statement=True)
        self.load_statement_parties()

    def load_statement_parties(self):
        state = getattr(self, "statement_state", None)
        if not state: return
        try: parties = self.client.parties()
        except Exception: parties = []
        self.party_rows = parties
        self.balance_party_map = {f'{p["name"]} | {p.get("account_number") or ""} | {p["kind"]}': p for p in parties}
        state["vars"]["party_box"]["values"] = list(self.balance_party_map)

    def refresh_statement_parties(self): self.load_statement_parties()

    def load_statement(self):
        state = getattr(self, "statement_state", None)
        if state and state.get("result"): self.run_balance_report(state)
