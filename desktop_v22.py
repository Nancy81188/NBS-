"""Version 2.2 screens: sales documents (debit / credit notes, duplicate, preview, import), and helpers."""
from __future__ import annotations

import os
import tempfile
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

NAVY, GOLD, LIGHT = "#071b2e", "#c9a96a", "#f3f6f8"
TREATMENTS = {"Taxable": "standard", "Taxable 11%": "standard", "Zero-rated": "zero_rated", "Zero-rated (export)": "zero_rated", "Exempt": "exempt", "Exempt (Art. 16-17)": "exempt", "Out of scope": "out_of_scope"}


class V22Mixin:
    # ------------------------------------------------------------ sales documents
    def sales_doc_type_changed(self):
        if not self.sales_edit_id: self.refresh_sales_number()
        self.sales_mode_label.config(text=f"NEW {self.sales_doc_type.get().upper()}", bg="#8B1E1E" if self.sales_doc_type.get() == "Credit Note" else GOLD,
                                     fg="white" if self.sales_doc_type.get() == "Credit Note" else NAVY)

    def open_sales_by_number(self):
        self.search_open_sales(); values = list(self.sales_open_box["values"])
        if len(values) == 1: self.sales_open_choice.set(values[0]); self.open_sales_invoice()
        elif not values: messagebox.showwarning("Sales Invoice", f"No invoice number matches {self.sales_open_choice.get()}")
        else: self.sales_open_box.event_generate("<Down>")

    def duplicate_sales_invoice(self):
        if not self.sales_edit_id: return messagebox.showwarning("Sales Invoice", "Open the invoice to duplicate first (Find No.)")
        source = self.sales_no.get(); self.sales_edit_id = None
        self.sales_date.set(datetime.now().strftime("%d-%m-%Y")); self.sales_amount_paid.set("0"); self.refresh_sales_number()
        self.sales_mode_label.config(text=f"COPY OF {source}", bg=GOLD, fg=NAVY)
        messagebox.showinfo("Sales Invoice", f"A copy of {source} is ready as {self.sales_no.get()}. Change what you need and press Save.")

    def current_sales_document(self):
        self.update_sales_totals(); calc = getattr(self, "sales_calculation", None)
        if not calc or not calc["lines"]: raise ValueError("Add at least one line")
        treatment = TREATMENTS.get(self.sales_treatment.get(), "standard")
        invoice = {"invoice_number": self.sales_no.get(), "invoice_date": self.sales_date.get(), "due_date": self.sales_due_date.get(), "party_name": self.sales_party.get(),
                   "currency": self.sales_currency.get(), "kind": "sale", "subtotal": calc["total_ht"], "vat": calc["vat"], "total": calc["grand_total"],
                   "invoice_discount_percent": self.sales_discount_percent.get(), "invoice_discount_amount": self.sales_discount_amount.get(),
                   "vat_treatment": treatment, "payment_method": self.sales_payment_method.get(), "amount_paid": self.sales_amount_paid.get(),
                   "doc_subtype": {"Credit Note": "credit_note", "Debit Note": "debit_note"}.get(self.sales_doc_type.get(), "invoice")}
        return invoice, calc["lines"]

    def sales_invoice_pdf(self, mode):
        from report_export import export_invoice_pdf
        try: invoice, lines = self.current_sales_document()
        except ValueError as exc: return messagebox.showwarning("Sales Invoice", str(exc))
        try: company = self.client.settings()
        except Exception: company = {}
        if mode == "pdf":
            path = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile=f"{invoice['invoice_number'] or 'invoice'}.pdf", filetypes=[("PDF", "*.pdf")])
            if not path: return
        else:
            handle = tempfile.NamedTemporaryFile(prefix=f"{invoice['invoice_number'] or 'invoice'}_", suffix=".pdf", delete=False); handle.close(); path = handle.name
        try: export_invoice_pdf(path, invoice, lines, company.get("company_logo") or None, company)
        except Exception as exc: return messagebox.showerror("Sales Invoice", f"The PDF could not be made: {exc}")
        if mode == "pdf": return messagebox.showinfo("Sales Invoice", f"Saved: {path}")
        try:
            if os.name == "nt": os.startfile(path, "print" if mode == "print" else "open")
            else: raise RuntimeError("Preview and printing open in the Windows application")
        except Exception as exc: messagebox.showinfo("Sales Invoice", f"The PDF is ready: {path}\n{exc}")

    def save_invoice_template(self, kind):
        from importer import write_invoice_template
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=f"Saber_{kind}_invoice_template.xlsx", filetypes=[("Excel", "*.xlsx")])
        if not path: return
        try: write_invoice_template(path, kind)
        except Exception as exc: return messagebox.showerror("Excel Template", str(exc))
        messagebox.showinfo("Excel Template", f"Template saved: {path}\nFill the blank Invoices sheet, then use Import Excel. Examples are on a separate sheet.")

    def import_sales_excel(self):
        import invoice_calc
        from importer import read_invoice_lines
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm")])
        if not path: return
        try: invoices = read_invoice_lines(path, "sales")
        except Exception as exc: return messagebox.showerror("Import Sales", f"The Excel file could not be read: {exc}")
        if not invoices: return messagebox.showwarning("Import Sales", "The Invoices sheet is empty. Fill its rows and try again.")
        if not messagebox.askyesno("Import Sales", f"Import {len(invoices)} invoice(s) as drafts (Review status)?"): return
        done = 0; errors = []
        for invoice in invoices:
            treatment = TREATMENTS.get(invoice["vat_treatment"].strip().capitalize(), TREATMENTS.get(invoice["vat_treatment"], "standard"))
            try:
                calc = invoice_calc.calculate(invoice["lines"], zero_vat=treatment != "standard")
                header = {"invoice_number": invoice["invoice_number"], "invoice_date": invoice["invoice_date"], "party_name": invoice["party_name"], "kind": "sales",
                          "currency": invoice["currency"], "status": "review", "source_file": Path(path).name, "vat_treatment": treatment, "gross_before_discount": calc["total"]}
                self.client.create_manual_invoice(header, [{k: v for k, v in l.items() if k != "net"} for l in calc["lines"]]); done += 1
            except Exception as exc: errors.append(f"{invoice['invoice_number']}: {exc}")
        (messagebox.showwarning if errors else messagebox.showinfo)("Import Sales", f"{done} invoice(s) imported as drafts." + ("\n" + "\n".join(errors[:12]) if errors else ""))
        self.load_sales_customer_list(); self.load_invoices(); self.load_journal(); self.load_trial()

    def import_sales_pdf(self):
        from pdf_import import read_invoice_pdf
        path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if not path: return
        data = read_invoice_pdf(path); self.new_sales_invoice(confirm=False)
        if data.get("invoice_date"): self.sales_date.set(data["invoice_date"])
        if data.get("party_name"): self.sales_party.set(data["party_name"])
        if data.get("currency"): self.sales_currency.set(data["currency"])
        first = self.sales_items[0]; first.update(description=f"As per {Path(path).name}", quantity=1, unit_price=data.get("subtotal") or data.get("total") or 0)
        self.recalculate_sales_item(first); self.sales_sheet.item(first["_iid"], values=self.sales_row_values(first)); self.update_sales_totals()
        messagebox.showinfo("Import PDF", f"{data.get('notes', '')}\nCheck the customer and the amount, then press Save.")

    # ------------------------------------------------------------ F2: the list that fits the field you are in
    def setup_context_f2(self):
        self.unbind_all("<F2>"); self.bind_all("<F2>", self.context_f2)
        parties = [(getattr(self, "sales_party_box", None), ("customer", "both")), (getattr(self, "purchase_form", {}).get("supplier_box"), ("supplier", "both")),
                   (getattr(self, "sd_party_box", None), None)]
        for form in getattr(self, "payment_forms", {}).values(): parties.append((form.get("party_box"), None))
        for state in (getattr(self, "statement_state", None),):
            if state: parties.append((state["vars"].get("party_box"), None))
        for box, kinds in parties:
            if box is not None: box._f2 = (lambda b=box, k=kinds: self.party_picker(b, k))
        items = [(getattr(self, "sales_sheet", None), self.sales_item_f2), (getattr(self, "stock_sheet", None) and self.stock_sheet.tree, self.stock_item_lookup),
                 (getattr(self, "sio_sheet", None) and self.sio_sheet.tree, lambda: self.item_picker(self.sio_pick))]
        if getattr(self, "purchase_form", None): items.append((self.purchase_form["items_sheet"].tree, self.purchase_item_lookup))
        for widget, action in items:
            if widget is not None: widget._f2 = action

    def context_f2(self, event=None):
        widget = self.focus_get()
        while widget is not None:
            action = getattr(widget, "_f2", None)
            if action: action(); return "break"
            widget = getattr(widget, "master", None)
        return self.open_active_account_lookup(event)

    def sales_item_f2(self):
        iid = self.sales_sheet.focus() or (self.sales_items[-1]["_iid"] if self.sales_items else None)
        if not iid: return
        def chosen(sku):
            item = self.sales_item_for(iid); product = self.item_by_code(sku)
            if not item or not product: return
            item.update(item_code=product["sku"], description=product["name"], unit=product.get("unit") or "", unit_price=product["sales_price"] or item.get("unit_price") or 0)
            self.recalculate_sales_item(item); self.sales_sheet.item(iid, values=self.sales_row_values(item)); self.update_sales_totals()
        self.item_picker(chosen)

    def party_picker(self, box, kinds=None):
        try: parties = self.client.parties()
        except Exception as exc: return messagebox.showerror("Customers / Suppliers", str(exc))
        parties = [p for p in parties if not kinds or p["kind"] in kinds]
        window = tk.Toplevel(self); window.title("Customers / Suppliers - F2"); window.geometry("700x420"); window.configure(bg=LIGHT); window.transient(self); window.grab_set()
        search = tk.StringVar(); entry = tk.Entry(window, textvariable=search, width=40); entry.pack(padx=10, pady=8); entry.focus_set()
        tree = ttk.Treeview(window, columns=("name", "account", "kind", "currency"), show="headings"); tree.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        for key, label, width in (("name", "Name", 300), ("account", "Account", 110), ("kind", "Type", 90), ("currency", "Currency", 70)): tree.heading(key, text=label); tree.column(key, width=width)
        def fill(*_a):
            tree.delete(*tree.get_children()); text = search.get().casefold()
            for p in parties:
                if not text or text in f'{p["name"]} {p.get("account_number") or ""} {p.get("tax_number") or ""}'.casefold():
                    tree.insert("", "end", iid=str(p["id"]), values=(p["name"], p.get("account_number") or "", p["kind"], p.get("currency") or ""))
        def choose(_e=None):
            if not tree.selection(): return
            party = next(p for p in parties if str(p["id"]) == tree.selection()[0]); window.destroy()
            label = next((v for v in box["values"] if str(v).split(" | ")[0] == party["name"]), party["name"])
            box.set(label); box.event_generate("<<ComboboxSelected>>")
        search.trace_add("write", fill); tree.bind("<Double-1>", choose); tree.bind("<Return>", choose); fill()
