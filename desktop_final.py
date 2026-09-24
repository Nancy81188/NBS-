"""Desktop screens added in version 1.12: payroll official reports, quarterly VAT return,
user expiry and permissions, legal-document alerts, and friendly error handling."""
from __future__ import annotations

import tkinter as tk
import traceback
from datetime import datetime, timedelta
from tkinter import filedialog, messagebox, ttk

from report_export import export_sections_excel, export_sections_pdf
import vat_return as vat_rules

NAVY, GOLD, LIGHT = "#071b2e", "#c9a96a", "#f3f6f8"
RED, AMBER, MUTED = "#8B1E1E", "#8a5a00", "#5f6b76"
MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]


def _user_date(value):
    text = str(value or "").strip()
    for pattern in ("%d-%m-%Y", "%d%m%Y", "%Y-%m-%d", "%Y%m%d"):
        try: return datetime.strptime(text, pattern)
        except ValueError: pass
    raise ValueError("Date must use DD-MM-YYYY")


def _display(value):
    text = str(value or "")
    return f"{text[8:10]}-{text[5:7]}-{text[:4]}" if len(text) == 10 and text[4] == "-" else text


def _fmt(value):
    if isinstance(value, bool) or value is None: return "" if value is None else str(value)
    if isinstance(value, (int, float)):
        return f"{value:,.0f}" if abs(value - round(value)) < 1e-9 else f"{value:,.2f}"
    return str(value)


class FinalFeaturesMixin:
    # ------------------------------------------------------------ errors and session
    def report_callback_exception(self, exc_type, exc, tb):
        """Show a clear message instead of a silent failure or a raw traceback."""
        traceback.print_exception(exc_type, exc, tb)
        if exc_type.__name__ == "SessionExpired": return
        messagebox.showerror("Saber Accounting", f"Something went wrong: {exc}\n\nYour saved data is safe. Please try again; if this continues, restart the application.")

    def session_ended(self):
        if getattr(self, "_session_notice", False): return
        self._session_notice = True
        def back_to_login():
            self.client = None; self.current_user = None; self.login_screen()
            messagebox.showinfo("Saber Accounting", "Your session has ended or your account is no longer active. Please sign in again.")
            self._session_notice = False
        self.after(0, back_to_login)

    def can_use(self, module):
        user = getattr(self, "current_user", None) or {}
        return user.get("role") == "admin" or bool((user.get("permissions") or {}).get(module, True))

    def status_bar(self):
        user = self.current_user or {}
        bar = tk.Frame(self, bg=NAVY); bar.pack(fill="x", side="bottom")
        expires = user.get("expires_at")
        validity = f"Account valid until {_display(expires)}" if expires else "Account without expiry"
        text = f"Signed in: {user.get('username', '')}   |   Role: {str(user.get('role', '')).title()}   |   {validity}"
        tk.Label(bar, text=text, bg=NAVY, fg="white", font=("Segoe UI", 8)).pack(side="left", padx=12, pady=3)
        if expires:
            try:
                days = (datetime.strptime(expires, "%Y-%m-%d").date() - datetime.now().date()).days
                if days <= 30: tk.Label(bar, text=f"Account expires in {days} day(s) - ask the administrator to renew", bg=NAVY, fg=GOLD, font=("Segoe UI", 8, "bold")).pack(side="left", padx=8)
            except ValueError: pass

    def scrollable_page(self, parent):
        """A notebook page with a vertical scrollbar for long forms."""
        outer = tk.Frame(parent, bg=LIGHT)
        canvas = tk.Canvas(outer, bg=LIGHT, highlightthickness=0); scroll = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas, bg=LIGHT); window = canvas.create_window((0, 0), window=inner, anchor="nw")
        inner.bind("<Configure>", lambda _event: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda event: canvas.itemconfigure(window, width=event.width))
        canvas.configure(yscrollcommand=scroll.set); canvas.pack(side="left", fill="both", expand=True); scroll.pack(side="right", fill="y")
        def wheel(event):
            if canvas.winfo_ismapped(): canvas.yview_scroll(int(-1 * (event.delta / 120)) if event.delta else (1 if getattr(event, "num", 0) == 5 else -1), "units")
        inner.bind("<Enter>", lambda _event: canvas.bind_all("<MouseWheel>", wheel)); inner.bind("<Leave>", lambda _event: canvas.unbind_all("<MouseWheel>"))
        return outer, inner

    # ------------------------------------------------------------ legal document alerts
    def show_document_alerts(self, startup=False):
        try: result = self.client.document_alerts(30)
        except Exception as exc:
            if not startup: messagebox.showerror("Legal Document Alerts", str(exc))
            return
        items = result.get("items", [])
        if hasattr(self, "alerts_button"):
            self.alerts_button.config(text=f"Document Alerts ({len(items)})" if items else "Document Alerts",
                                      bg=RED if result.get("expired") else GOLD if items else NAVY, fg="white" if result.get("expired") or not items else NAVY)
        if startup and not items: return
        window = tk.Toplevel(self); window.title("Legal Document Alerts"); window.configure(bg=LIGHT); window.geometry("900x420"); window.transient(self)
        headline = (f"{result.get('expired', 0)} expired and {result.get('expiring', 0)} expiring within {result.get('days', 30)} days"
                    if items else "No expired or expiring legal documents")
        tk.Label(window, text=headline, bg=LIGHT, fg=RED if result.get("expired") else NAVY, font=("Segoe UI", 12, "bold")).pack(pady=(12, 4))
        tk.Label(window, text="Renew these documents and upload the new copy from Customers / Suppliers > Legal Documents.", bg=LIGHT, fg=MUTED).pack()
        frame = tk.Frame(window, bg=LIGHT); frame.pack(fill="both", expand=True, padx=10, pady=10)
        columns = (("status", "Status", 110), ("party", "Customer / Supplier", 230), ("type", "Document", 190), ("expiry", "Expiry Date", 100), ("days", "Days", 70), ("file", "File", 170))
        tree = ttk.Treeview(frame, columns=[c[0] for c in columns], show="headings")
        for key, label, width in columns: tree.heading(key, text=label); tree.column(key, width=width, anchor="w")
        tree.tag_configure("expired", foreground=RED); tree.tag_configure("soon", foreground=AMBER)
        scroll = ttk.Scrollbar(frame, orient="vertical", command=tree.yview); tree.configure(yscrollcommand=scroll.set)
        tree.pack(side="left", fill="both", expand=True); scroll.pack(side="right", fill="y")
        for row in items:
            days = row["days_remaining"]
            tree.insert("", "end", values=(row["status"].title(), row["party_name"], row["document_type"], row["expiry_date"],
                "expired" if days < 0 else days, row["file_name"]), tags=("expired" if days < 0 else "soon",))
        self.action_button(window, "Close", window.destroy).pack(pady=(0, 10))

    # ------------------------------------------------------------ payroll official reports
    def build_payroll_reports_page(self, page):
        controls = tk.Frame(page, bg=LIGHT); controls.pack(fill="x", padx=10, pady=8)
        now = datetime.now()
        self.pr_report = tk.StringVar(value="R10 - Quarterly withholding"); self.pr_period_type = tk.StringVar(value="Quarterly")
        self.pr_year = tk.StringVar(value=str(getattr(self, "current_fiscal_year", now.year))); self.pr_index = tk.StringVar(value=f"Q{(now.month - 1) // 3 + 1}")
        self.pr_group = tk.StringVar(value="Employees and Managers (separate)"); self.pr_drafts = tk.BooleanVar(value=False)
        reports = ["R10 - Quarterly withholding", "R5 - Annual employer declaration", "R6 - Individual annual statement"]
        tk.Label(controls, text="Report", bg=LIGHT).grid(row=0, column=0, padx=4, sticky="w")
        ttk.Combobox(controls, textvariable=self.pr_report, values=reports, state="readonly", width=31).grid(row=0, column=1, padx=4)
        tk.Label(controls, text="Period", bg=LIGHT).grid(row=0, column=2, padx=4, sticky="w")
        period_box = ttk.Combobox(controls, textvariable=self.pr_period_type, values=["Monthly", "Quarterly", "Yearly"], state="readonly", width=11); period_box.grid(row=0, column=3, padx=4)
        tk.Label(controls, text="Year", bg=LIGHT).grid(row=0, column=4, padx=4, sticky="w")
        tk.Entry(controls, textvariable=self.pr_year, width=7).grid(row=0, column=5, padx=4)
        self.pr_index_box = ttk.Combobox(controls, textvariable=self.pr_index, state="readonly", width=11); self.pr_index_box.grid(row=0, column=6, padx=4)
        tk.Label(controls, text="Group", bg=LIGHT).grid(row=1, column=0, padx=4, pady=6, sticky="w")
        ttk.Combobox(controls, textvariable=self.pr_group, values=["Employees and Managers (separate)", "Employees only", "Managers only"], state="readonly", width=31).grid(row=1, column=1, padx=4, pady=6)
        tk.Checkbutton(controls, text="Include draft payroll (preview only)", variable=self.pr_drafts, bg=LIGHT).grid(row=1, column=2, columnspan=3, sticky="w", padx=4)
        buttons = tk.Frame(controls, bg=LIGHT); buttons.grid(row=1, column=5, columnspan=4, sticky="e")
        tk.Button(buttons, text="Generate", command=self.generate_payroll_report, bg=GOLD, fg=NAVY, border=0, padx=16, pady=6, font=("Segoe UI", 9, "bold")).pack(side="left", padx=3)
        self.action_button(buttons, "Export Excel", lambda: self.export_payroll_report("xlsx")).pack(side="left", padx=3)
        self.action_button(buttons, "Export PDF", lambda: self.export_payroll_report("pdf")).pack(side="left", padx=3)
        def refresh_index(*_args):
            kind = self.pr_period_type.get()
            if kind == "Monthly":
                self.pr_index_box.config(values=MONTHS, state="readonly")
                if self.pr_index.get() not in MONTHS: self.pr_index.set(MONTHS[now.month - 1])
            elif kind == "Quarterly":
                self.pr_index_box.config(values=["Q1", "Q2", "Q3", "Q4"], state="readonly")
                if self.pr_index.get() not in ("Q1", "Q2", "Q3", "Q4"): self.pr_index.set(f"Q{(now.month - 1) // 3 + 1}")
            else:
                self.pr_index_box.config(values=["Full year"], state="disabled"); self.pr_index.set("Full year")
        period_box.bind("<<ComboboxSelected>>", refresh_index); refresh_index()
        self.pr_info = tk.Label(page, text="Choose a report and period, then press Generate. Official figures use posted payroll, converted to LBP.", bg=LIGHT, fg=MUTED, anchor="w", justify="left")
        self.pr_info.pack(fill="x", padx=12)
        self.pr_tree = self.report_viewer(page)
        self.payroll_report_result = None

    def report_viewer(self, parent):
        frame = tk.Frame(parent, bg=LIGHT); frame.pack(fill="both", expand=True, padx=10, pady=8)
        columns = [f"c{i}" for i in range(20)]
        tree = ttk.Treeview(frame, columns=columns, show="", selectmode="browse")
        for index, column in enumerate(columns): tree.column(column, width=190 if index == 1 else 115, anchor="w", stretch=False)
        tree.tag_configure("section", background=NAVY, foreground="white", font=("Segoe UI", 9, "bold"))
        tree.tag_configure("header", background="#dfe6ee", foreground=NAVY, font=("Segoe UI", 8, "bold"))
        tree.tag_configure("total", background="#e8edf2", font=("Segoe UI", 9, "bold"))
        yscroll = ttk.Scrollbar(frame, orient="vertical", command=tree.yview); xscroll = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        tree.grid(row=0, column=0, sticky="nsew"); yscroll.grid(row=0, column=1, sticky="ns"); xscroll.grid(row=1, column=0, sticky="ew")
        frame.grid_rowconfigure(0, weight=1); frame.grid_columnconfigure(0, weight=1)
        return tree

    def show_sections(self, tree, sections):
        tree.delete(*tree.get_children())
        for section in sections:
            tree.insert("", "end", values=[section["heading"]], tags=("section",))
            tree.insert("", "end", values=section["headers"], tags=("header",))
            totals = set(section.get("total_rows") or [])
            for index, row in enumerate(section["rows"]):
                tree.insert("", "end", values=[_fmt(value) for value in row], tags=("total",) if index in totals else ())
            tree.insert("", "end", values=[""])

    def payroll_report_parameters(self):
        report = self.pr_report.get().split(" ", 1)[0]
        period = self.pr_period_type.get().lower()
        try: year = int(self.pr_year.get().strip())
        except ValueError: raise ValueError("Enter the year as four digits, for example 2025")
        if period == "monthly": index = MONTHS.index(self.pr_index.get()) + 1
        elif period == "quarterly": index = int(self.pr_index.get().lstrip("Q"))
        else: index = 1
        group = {"Employees only": "employee", "Managers only": "manager"}.get(self.pr_group.get(), "both")
        return report, period, year, index, group

    def generate_payroll_report(self):
        try:
            report, period, year, index, group = self.payroll_report_parameters()
            result = self.client.payroll_report(report, period, year, index, group, self.pr_drafts.get())
        except Exception as exc: return messagebox.showerror("Payroll Reports", str(exc))
        self.payroll_report_result = result
        self.show_sections(self.pr_tree, result["sections"])
        note = f"{result['title']}  |  {result['period_label']}  |  {result['record_count']} payroll record(s)"
        if not result["record_count"]:
            note += "  |  No posted payroll in this period. Post payroll records, or tick 'Include draft payroll' to preview."
        self.pr_info.config(text=note, fg=RED if not result["record_count"] else NAVY)

    def export_payroll_report(self, format_name):
        result = getattr(self, "payroll_report_result", None)
        if not result: return messagebox.showwarning("Payroll Reports", "Generate the report first")
        name = f"{result['report']}_{result['period_label'].split(' (')[0].replace(' ', '_')}"
        self.save_sections(result["title"], result["meta"], result["sections"], name, format_name)

    def save_sections(self, title, meta, sections, name, format_name):
        extension = ".xlsx" if format_name == "xlsx" else ".pdf"
        path = filedialog.asksaveasfilename(defaultextension=extension, initialfile=name + extension,
            filetypes=[("Excel workbook", "*.xlsx")] if format_name == "xlsx" else [("PDF document", "*.pdf")])
        if not path: return
        try: (export_sections_excel if format_name == "xlsx" else export_sections_pdf)(path, title, meta, sections)
        except PermissionError: return messagebox.showerror(title, "The file could not be saved. Close it if it is open in Excel or a PDF viewer, then try again.")
        except Exception as exc: return messagebox.showerror(title, f"The file could not be saved: {exc}")
        messagebox.showinfo(title, f"Saved successfully:\n{path}")

    def build_payroll_periods_panel(self, parent, row):
        frame = tk.LabelFrame(parent, text="Effective periods (rates and ceilings apply From Date to To Date)", bg=LIGHT, padx=6, pady=4)
        frame.grid(row=row, column=0, columnspan=6, padx=10, pady=6, sticky="ew")
        columns = (("from", "Date From", 90), ("to", "Date To", 90), ("emp", "Employee Rate", 90), ("emp_c", "Employee Ceiling", 115), ("med", "Medical Rate", 85),
                   ("med_c", "Medical Ceiling", 110), ("fam", "Family Rate", 80), ("fam_c", "Family Ceiling", 105), ("eos", "EOS Rate", 70), ("eos_c", "EOS Ceiling", 100))
        self.payroll_periods_tree = ttk.Treeview(frame, columns=[c[0] for c in columns], show="headings", height=4)
        for key, label, width in columns: self.payroll_periods_tree.heading(key, text=label); self.payroll_periods_tree.column(key, width=width, anchor="w")
        self.payroll_periods_tree.pack(fill="x")
        self.payroll_periods_tree.bind("<Double-1>", lambda _event: self.load_selected_payroll_period())
        tk.Label(frame, text="Double-click a period to load it. Saving a new Date From automatically ends the previous period the day before.", bg=LIGHT, fg=MUTED).pack(anchor="w")
        self.load_payroll_periods()

    def load_payroll_periods(self):
        if not hasattr(self, "payroll_periods_tree"): return
        try: rows = self.client.payroll_settings_list()
        except Exception: return
        self.payroll_period_rows = {row["date_from"]: row for row in rows}
        self.payroll_periods_tree.delete(*self.payroll_periods_tree.get_children())
        def rate(value): return f"{float(value) * 100:g}%"
        def ceiling(value): return "No ceiling" if not float(value or 0) else f"{float(value):,.0f}"
        for row in rows:
            self.payroll_periods_tree.insert("", "end", iid=row["date_from"], values=(_display(row["date_from"]), _display(row.get("date_to")) or "Open",
                rate(row["employee_nssf_rate"]), ceiling(row["employee_ceiling"]), rate(row["medical_rate"]), ceiling(row["medical_ceiling"]),
                rate(row["family_rate"]), ceiling(row["family_ceiling"]), rate(row["end_service_rate"]), ceiling(row["end_service_ceiling"])))

    def load_selected_payroll_period(self):
        selected = self.payroll_periods_tree.selection()
        if not selected: return
        self.payroll_period.set(_display(selected[0])); self.load_payroll_settings()

    # ------------------------------------------------------------ quarterly VAT return
    def build_vat_return(self):
        page = self.vat_tab; now = datetime.now()
        controls = tk.Frame(page, bg=LIGHT); controls.pack(fill="x", padx=10, pady=8)
        self.vat_year = tk.StringVar(value=str(getattr(self, "current_fiscal_year", now.year))); self.vat_quarter = tk.StringVar(value=f"Q{(now.month - 1) // 3 + 1}")
        self.vat_currency = tk.StringVar(value="All Currencies"); self.vat_include_review = tk.BooleanVar(value=False); self.vat_credit_override = tk.StringVar()
        tk.Label(controls, text="Year", bg=LIGHT).pack(side="left"); tk.Entry(controls, textvariable=self.vat_year, width=7).pack(side="left", padx=4)
        tk.Label(controls, text="Quarter", bg=LIGHT).pack(side="left", padx=(6, 0))
        ttk.Combobox(controls, textvariable=self.vat_quarter, values=["Q1", "Q2", "Q3", "Q4"], state="readonly", width=5).pack(side="left", padx=4)
        tk.Label(controls, text="Currency", bg=LIGHT).pack(side="left", padx=(6, 0))
        ttk.Combobox(controls, textvariable=self.vat_currency, values=["All Currencies", "USD", "EUR", "LBP", "AED"], state="readonly", width=13).pack(side="left", padx=4)
        tk.Checkbutton(controls, text="Include Review documents", variable=self.vat_include_review, bg=LIGHT).pack(side="left", padx=6)
        tk.Label(controls, text="Credit B/F (LBP, optional)", bg=LIGHT).pack(side="left", padx=(6, 0))
        tk.Entry(controls, textvariable=self.vat_credit_override, width=14).pack(side="left", padx=4)
        tk.Button(controls, text="Generate", command=self.load_vat_return, bg=GOLD, fg=NAVY, border=0, padx=16, pady=6, font=("Segoe UI", 9, "bold")).pack(side="left", padx=4)
        actions = tk.Frame(page, bg=LIGHT); actions.pack(fill="x", padx=10)
        self.action_button(actions, "Save Return (lock quarter)", self.save_vat_return).pack(side="left", padx=3)
        if (self.current_user or {}).get("role") == "admin":
            tk.Button(actions, text="Reopen Saved Return", command=self.reopen_vat_return, bg=RED, fg="white", border=0, padx=14, pady=7).pack(side="left", padx=3)
        self.action_button(actions, "Export Excel", lambda: self.export_vat_return("xlsx")).pack(side="left", padx=3)
        self.action_button(actions, "Export PDF", lambda: self.export_vat_return("pdf")).pack(side="left", padx=3)
        self.vat_headline = tk.Label(page, text="Choose the year and quarter, then press Generate.", bg=LIGHT, fg=NAVY, font=("Segoe UI", 11, "bold"), anchor="w", justify="left")
        self.vat_headline.pack(fill="x", padx=12, pady=(8, 0))
        self.vat_note = tk.Label(page, text="", bg=LIGHT, fg=MUTED, anchor="w", justify="left"); self.vat_note.pack(fill="x", padx=12)
        nested = ttk.Notebook(page); nested.pack(fill="both", expand=True, padx=10, pady=8)
        summary = tk.Frame(nested, bg=LIGHT); documents = tk.Frame(nested, bg=LIGHT); adjustments = tk.Frame(nested, bg=LIGHT); history = tk.Frame(nested, bg=LIGHT)
        nested.add(summary, text="VAT Return"); nested.add(documents, text="Supporting Documents"); nested.add(adjustments, text="Manual Adjustments"); nested.add(history, text="Saved Returns")
        self.vat_summary_tree = self.report_viewer(summary)
        self.vat_documents_tree = self.table(documents, [("date", "Date", 90), ("number", "Document", 120), ("party", "Customer / Supplier", 200), ("category", "Category", 140),
            ("deductible", "Deductible", 80), ("currency", "Currency", 70), ("base", "Base", 110), ("vat", "VAT", 100), ("rate", "LBP Rate", 90), ("vat_lbp", "VAT (LBP)", 120), ("status", "Status", 75)])
        form = tk.Frame(adjustments, bg=LIGHT); form.pack(fill="x", padx=10, pady=8)
        self.vat_adj_type = tk.StringVar(value="Output VAT"); self.vat_adj_currency = tk.StringVar(value="LBP"); self.vat_adj_amount = tk.StringVar(); self.vat_adj_reason = tk.StringVar()
        ttk.Combobox(form, textvariable=self.vat_adj_type, values=["Output VAT", "Deductible VAT", "Non-deductible VAT"], state="readonly", width=18).pack(side="left", padx=3)
        ttk.Combobox(form, textvariable=self.vat_adj_currency, values=["LBP", "USD", "EUR", "AED"], state="readonly", width=6).pack(side="left", padx=3)
        tk.Label(form, text="Amount (+/-)", bg=LIGHT).pack(side="left"); tk.Entry(form, textvariable=self.vat_adj_amount, width=14).pack(side="left", padx=3)
        tk.Label(form, text="Reason", bg=LIGHT).pack(side="left"); tk.Entry(form, textvariable=self.vat_adj_reason, width=38).pack(side="left", padx=3)
        self.action_button(form, "Add Adjustment", self.add_vat_adjustment).pack(side="left", padx=3)
        tk.Button(form, text="Delete Selected", command=self.delete_vat_adjustment, bg=RED, fg="white", border=0, padx=12, pady=7).pack(side="left", padx=3)
        self.vat_adjustments_tree = self.table(adjustments, [("type", "Type", 190), ("currency", "Currency", 70), ("amount", "Amount", 110), ("amount_lbp", "Amount (LBP)", 130),
            ("reason", "Reason", 300), ("by", "Entered by", 100), ("at", "Entered at", 140)])
        self.vat_history_tree = self.table(history, [("period", "Quarter", 90), ("net", "Net VAT (LBP)", 140), ("bf", "Credit B/F", 130), ("payable", "Payable", 130),
            ("cf", "Credit C/F", 130), ("by", "Saved by", 100), ("at", "Saved at", 150)])
        self.vat_return_result = None

    def vat_parameters(self):
        try: year = int(self.vat_year.get().strip())
        except ValueError: raise ValueError("Enter the year as four digits, for example 2025")
        quarter = int(self.vat_quarter.get().lstrip("Q"))
        currency = None if self.vat_currency.get() == "All Currencies" else self.vat_currency.get()
        credit = self.vat_credit_override.get().strip().replace(",", "") or None
        if credit is not None:
            try: float(credit)
            except ValueError: raise ValueError("Credit brought forward must be a number in LBP, or left empty")
        return year, quarter, currency, credit

    def load_vat_return(self):
        if not hasattr(self, "vat_summary_tree"): return
        try:
            year, quarter, currency, credit = self.vat_parameters()
            result = self.client.vat_return(year, quarter, currency, self.vat_include_review.get(), credit)
            history = self.client.vat_returns()
        except Exception as exc: return messagebox.showerror("Quarterly VAT", str(exc))
        self.vat_return_result = result
        title, meta, sections = vat_rules.export_sections(result)
        self.show_sections(self.vat_summary_tree, sections[:-2] if result["adjustments"] else sections[:-1])
        payable = result["payable_lbp"]; credit_cf = result["credit_carried_forward_lbp"]
        outcome = f"VAT PAYABLE: {payable:,.0f} LBP" if payable else f"CREDIT CARRIED FORWARD: {credit_cf:,.0f} LBP" if credit_cf else "NIL RETURN: 0 LBP"
        self.vat_headline.config(text=f"Q{quarter} {year}  |  {_display(result['date_from'])} to {_display(result['date_to'])}  |  {outcome}  |  Status: {result['status'].title()}",
                                 fg=RED if result["changed_since_saved"] else NAVY)
        notes = [f"Credit brought forward: {result['credit_brought_forward_lbp']:,.0f} LBP ({result['credit_source']})"]
        if result["review_excluded"]: notes.append(f"{result['review_excluded']} document(s) in Review status are not included")
        if result["skipped"]: notes.append(f"{len(result['skipped'])} document(s) have an unreadable date: {', '.join(map(str, result['skipped'][:5]))}")
        if result["changed_since_saved"]: notes.append("Documents changed after this return was saved - review and save again")
        if result["currency_filter"] != "All": notes.append("Currency filter active: payable/credit totals cover this currency only")
        self.vat_note.config(text="   |   ".join(notes))
        self.vat_documents_tree.delete(*self.vat_documents_tree.get_children())
        for d in result["documents"]:
            self.vat_documents_tree.insert("", "end", values=(_display(d["date"]), d["number"], d["party"], vat_rules.CATEGORIES[d["category"]], "Yes" if d["recoverable"] else "No",
                d["currency"], _fmt(d["base"]), _fmt(d["vat"]), _fmt(d["lbp_rate"]), _fmt(d["vat_lbp"]), d["status"]))
        self.vat_adjustment_rows = {str(a["id"]): a for a in result["adjustments"]}
        self.vat_adjustments_tree.delete(*self.vat_adjustments_tree.get_children())
        for a in result["adjustments"]:
            self.vat_adjustments_tree.insert("", "end", iid=str(a["id"]), values=(vat_rules.ADJUSTMENT_TYPES[a["adjustment_type"]], a["currency"], _fmt(float(a["amount"])),
                _fmt(a.get("amount_lbp", "")), a["reason"], a.get("created_by_name") or "", str(a["created_at"])[:16].replace("T", " ")))
        self.vat_history_tree.delete(*self.vat_history_tree.get_children())
        for row in history:
            self.vat_history_tree.insert("", "end", values=(f"Q{row['quarter']} {row['year']}", _fmt(float(row["net_lbp"])), _fmt(float(row["credit_brought_forward_lbp"])),
                _fmt(float(row["payable_lbp"])), _fmt(float(row["credit_carried_forward_lbp"])), row.get("saved_by_name") or "", str(row["saved_at"])[:16].replace("T", " ")))

    def add_vat_adjustment(self):
        try:
            year, quarter, _currency, _credit = self.vat_parameters()
            kind = {"Output VAT": "output", "Deductible VAT": "input", "Non-deductible VAT": "non_deductible"}[self.vat_adj_type.get()]
            amount = self.vat_adj_amount.get().strip().replace(",", "")
            if not amount: raise ValueError("Enter the adjustment amount (use a minus sign to reduce VAT)")
            self.client.add_vat_adjustment({"year": year, "quarter": quarter, "adjustment_type": kind, "currency": self.vat_adj_currency.get(),
                                            "amount": amount, "reason": self.vat_adj_reason.get().strip()})
        except Exception as exc: return messagebox.showerror("VAT Adjustment", str(exc))
        self.vat_adj_amount.set(""); self.vat_adj_reason.set(""); self.load_vat_return()

    def delete_vat_adjustment(self):
        selected = self.vat_adjustments_tree.selection()
        if not selected: return messagebox.showwarning("VAT Adjustment", "Select an adjustment first")
        record = getattr(self, "vat_adjustment_rows", {}).get(selected[0])
        if not record or not messagebox.askyesno("VAT Adjustment", f"Delete the adjustment '{record['reason']}'?"): return
        try: self.client.delete_vat_adjustment(record["id"])
        except Exception as exc: return messagebox.showerror("VAT Adjustment", str(exc))
        self.load_vat_return()

    def save_vat_return(self):
        try: year, quarter, _currency, credit = self.vat_parameters()
        except Exception as exc: return messagebox.showerror("Quarterly VAT", str(exc))
        if not messagebox.askyesno("Save VAT Return", f"Save the Q{quarter} {year} return for all currencies? Adjustments for this quarter will be locked, "
                                   "and the credit carried forward will be used by the next quarter."): return
        try: result = self.client.save_vat_return(year, quarter, credit)
        except Exception as exc: return messagebox.showerror("Quarterly VAT", str(exc))
        self.vat_currency.set("All Currencies"); self.load_vat_return()
        messagebox.showinfo("Quarterly VAT", f"Q{quarter} {year} saved. Payable: {result['payable_lbp']:,.0f} LBP   Credit carried forward: {result['credit_carried_forward_lbp']:,.0f} LBP")

    def reopen_vat_return(self):
        try: year, quarter, _currency, _credit = self.vat_parameters()
        except Exception as exc: return messagebox.showerror("Quarterly VAT", str(exc))
        if not messagebox.askyesno("Reopen VAT Return", f"Reopen Q{quarter} {year}? It will need to be saved again after changes."): return
        try: self.client.reopen_vat_return(year, quarter)
        except Exception as exc: return messagebox.showerror("Quarterly VAT", str(exc))
        self.load_vat_return()

    def export_vat_return(self, format_name):
        result = getattr(self, "vat_return_result", None)
        if not result: return messagebox.showwarning("Quarterly VAT", "Generate the return first")
        title, meta, sections = vat_rules.export_sections(result)
        self.save_sections(title, meta, sections, f"VAT_Return_Q{result['quarter']}_{result['year']}", format_name)

    def toggle_selected_invoice_vat(self):
        selected = self.invoice_tree.selection()
        if not selected: return messagebox.showwarning("VAT Deductibility", "Select a purchase or expense row first")
        row = getattr(self, "invoice_rows", {}).get(selected[0])
        if not row: return
        if row.get("kind") != "purchase": return messagebox.showwarning("VAT Deductibility", "Only purchase and expense VAT can be marked non-deductible")
        recoverable = bool(row.get("vat_recoverable", 1))
        action = "NON-DEDUCTIBLE (the VAT becomes part of the cost)" if recoverable else "DEDUCTIBLE again"
        if not messagebox.askyesno("VAT Deductibility", f"Mark the VAT of {row['invoice_number']} ({float(row.get('vat') or 0):,.2f} {row['currency']}) as {action}?"): return
        try: self.client.set_vat_recoverable("invoice", row["id"], not recoverable)
        except Exception as exc: return messagebox.showerror("VAT Deductibility", str(exc))
        self.load_invoices(); self.load_journal(); self.load_trial()

    # ------------------------------------------------------------ users: expiry and permissions
    def build_users_page(self, users):
        form = tk.Frame(users, bg=LIGHT); form.pack(fill="x", padx=10, pady=10)
        self.user_name = tk.StringVar(); self.user_password = tk.StringVar(); self.user_role = tk.StringVar(value="accountant"); self.user_language = tk.StringVar(value="en")
        self.user_expiry = tk.StringVar(value=(datetime.now() + timedelta(days=365)).strftime("%d-%m-%Y"))
        self.user_active = tk.BooleanVar(value=True); self.user_payroll = tk.BooleanVar(value=True); self.user_vat = tk.BooleanVar(value=True); self.edit_user_id = None
        fields = (("Username", self.user_name, 16, ""), ("Password", self.user_password, 14, "*"))
        for column, (label, var, width, show) in enumerate(fields):
            tk.Label(form, text=label, bg=LIGHT).grid(row=0, column=column * 2, padx=4, sticky="w")
            tk.Entry(form, textvariable=var, width=width, show=show).grid(row=0, column=column * 2 + 1, padx=4)
        tk.Label(form, text="Role", bg=LIGHT).grid(row=0, column=4, padx=4, sticky="w")
        ttk.Combobox(form, textvariable=self.user_role, values=["admin", "accountant", "viewer"], state="readonly", width=11).grid(row=0, column=5, padx=4)
        tk.Label(form, text="Language", bg=LIGHT).grid(row=0, column=6, padx=4, sticky="w")
        ttk.Combobox(form, textvariable=self.user_language, values=["en", "ar", "fr"], state="readonly", width=5).grid(row=0, column=7, padx=4)
        tk.Label(form, text="Valid until", bg=LIGHT).grid(row=1, column=0, padx=4, pady=6, sticky="w")
        self.date_entry(form, self.user_expiry, 12).grid(row=1, column=1, padx=4, pady=6, sticky="w")
        tk.Checkbutton(form, text="Active", variable=self.user_active, bg=LIGHT).grid(row=1, column=2, sticky="w")
        tk.Checkbutton(form, text="Payroll access", variable=self.user_payroll, bg=LIGHT).grid(row=1, column=3, sticky="w")
        tk.Checkbutton(form, text="VAT access", variable=self.user_vat, bg=LIGHT).grid(row=1, column=4, sticky="w")
        buttons = tk.Frame(form, bg=LIGHT); buttons.grid(row=1, column=5, columnspan=4, sticky="w")
        self.action_button(buttons, "Save User", self.add_user).pack(side="left", padx=3)
        self.action_button(buttons, "Renew 1 Year", self.renew_selected_user).pack(side="left", padx=3)
        self.action_button(buttons, "New / Clear", self.clear_user_form).pack(side="left", padx=3)
        tk.Label(users, text="Roles: Admin = everything; Accountant = enter and post; Viewer = read only. New non-admin users are valid for 1 year; "
                 "leave 'Valid until' empty for no expiry. Payroll and VAT access can be removed per user.", bg=LIGHT, fg=MUTED, wraplength=1050, justify="left").pack(fill="x", padx=12)
        self.users_tree = self.table(users, [("id", "ID", 50), ("username", "Username", 160), ("role", "Role", 95), ("language", "Language", 70), ("status", "Status", 80),
            ("expires", "Valid Until", 95), ("days", "Days Left", 75), ("payroll", "Payroll", 65), ("vat", "VAT", 55)])
        self.users_tree.tag_configure("expired", foreground=RED); self.users_tree.tag_configure("soon", foreground=AMBER)
        self.users_tree.bind("<Double-1>", lambda _event: self.edit_selected_user())

    def clear_user_form(self):
        self.edit_user_id = None; self.user_name.set(""); self.user_password.set(""); self.user_role.set("accountant"); self.user_active.set(True)
        self.user_payroll.set(True); self.user_vat.set(True)
        self.user_expiry.set((datetime.now() + timedelta(days=365)).strftime("%d-%m-%Y"))

    def fill_users_tree(self, users):
        self.user_rows = {str(row["id"]): row for row in users}
        self.users_tree.delete(*self.users_tree.get_children())
        for row in users:
            days = row.get("days_remaining"); tag = "expired" if row.get("status") == "expired" else "soon" if days is not None and days <= 30 else ""
            permissions = row.get("permissions") or {}
            self.users_tree.insert("", "end", iid=str(row["id"]), values=(row["id"], row["username"], row["role"], row["language"], str(row.get("status", "")).title(),
                _display(row.get("expires_at")) or "No expiry", "" if days is None else days, "Yes" if row["role"] == "admin" or permissions.get("payroll", True) else "No",
                "Yes" if row["role"] == "admin" or permissions.get("vat", True) else "No"), tags=(tag,) if tag else ())

    def user_payload(self):
        expiry = self.user_expiry.get().strip()
        if expiry: _user_date(expiry)
        return {"id": self.edit_user_id, "username": self.user_name.get().strip(), "password": self.user_password.get(), "role": self.user_role.get(),
                "language": self.user_language.get(), "active": self.user_active.get(), "expires_at": expiry,
                "permissions": {"payroll": self.user_payroll.get(), "vat": self.user_vat.get()}}

    def add_user(self):
        try: payload = self.user_payload(); saved = self.client.save_user(payload)
        except Exception as exc: return messagebox.showerror("Users", str(exc))
        self.clear_user_form(); self.load_settings_pages()
        messagebox.showinfo("Users", f"User {saved['username']} saved. Valid until: {_display(saved.get('expires_at')) or 'no expiry'}")

    def renew_selected_user(self):
        selected = self.users_tree.selection()
        if not selected: return messagebox.showwarning("Users", "Select a user to renew")
        row = self.user_rows.get(selected[0])
        try:
            saved = self.client.save_user({"id": row["id"], "username": row["username"], "role": row["role"], "language": row["language"],
                                           "active": True, "permissions": row.get("permissions") or {}, "renew": True})
        except Exception as exc: return messagebox.showerror("Users", str(exc))
        self.load_settings_pages(); messagebox.showinfo("Users", f"{saved['username']} renewed until {_display(saved.get('expires_at'))}")

    def edit_selected_user(self):
        selected = self.users_tree.selection()
        if not selected: return
        row = self.user_rows.get(selected[0])
        if not row: return
        self.edit_user_id = row["id"]; self.user_name.set(row["username"]); self.user_role.set(row["role"]); self.user_language.set(row["language"])
        self.user_password.set(""); self.user_active.set(bool(row["active"])); self.user_expiry.set(_display(row.get("expires_at")))
        permissions = row.get("permissions") or {}; self.user_payroll.set(permissions.get("payroll", True)); self.user_vat.set(permissions.get("vat", True))
