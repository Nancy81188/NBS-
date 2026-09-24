"""Departments, projects and budgets (version 1.15)."""
from __future__ import annotations

import tkinter as tk
from datetime import datetime
from tkinter import messagebox, ttk

from desktop_brains import EditableSheet

NAVY, GOLD, LIGHT = "#071b2e", "#c9a96a", "#f3f6f8"
RED, MUTED = "#8B1E1E", "#5f6b76"
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
NONE = "(none)"


def _num(value):
    try: return float(str(value or 0).replace(",", ""))
    except ValueError: return None


class DimensionsMixin:
    # ------------------------------------------------------------ shared lists
    def dimension_lists(self, refresh=False):
        if refresh or not getattr(self, "_dimensions", None):
            try: departments = self.client.departments(); projects = self.client.projects()
            except Exception: departments, projects = [], []
            self._dimensions = {"departments": departments, "projects": projects}
        return self._dimensions

    def department_choices(self, include_all=False):
        items = [f'{d["code"]} - {d["name"]}' for d in self.dimension_lists()["departments"] if d["active"]]
        return (["All"] if include_all else [NONE]) + items

    def project_choices(self, include_all=False):
        items = [f'{p["code"]} - {p["name"]}' for p in self.dimension_lists()["projects"] if p["active"] and p["status"] != "cancelled"]
        return (["All"] if include_all else [NONE]) + items

    @staticmethod
    def dimension_code(value):
        text = str(value or "").strip()
        return "" if text in ("", NONE, "All") else text.split(" - ", 1)[0].strip()

    def dimension_selectors(self, parent, department_var, project_var, include_all=False):
        tk.Label(parent, text="Department", bg=LIGHT).pack(side="left")
        department = ttk.Combobox(parent, textvariable=department_var, values=self.department_choices(include_all), state="readonly", width=20); department.pack(side="left", padx=(4, 10))
        tk.Label(parent, text="Project", bg=LIGHT).pack(side="left")
        project = ttk.Combobox(parent, textvariable=project_var, values=self.project_choices(include_all), state="readonly", width=24); project.pack(side="left", padx=(4, 10))
        def refresh(_event=None):
            self.dimension_lists(refresh=True); department["values"] = self.department_choices(include_all); project["values"] = self.project_choices(include_all)
        department.bind("<Button-1>", refresh, add="+"); project.bind("<Button-1>", refresh, add="+")
        if not department_var.get(): department_var.set("All" if include_all else NONE)
        if not project_var.get(): project_var.set("All" if include_all else NONE)
        return department, project

    # ------------------------------------------------------------ settings pages
    def build_dimensions_pages(self, nested):
        departments = tk.Frame(nested, bg=LIGHT); projects = tk.Frame(nested, bg=LIGHT)
        nested.add(departments, text="Departments"); nested.add(projects, text="Projects")
        self.dep_id = None; self.dep_code = tk.StringVar(); self.dep_name = tk.StringVar(); self.dep_active = tk.BooleanVar(value=True)
        form = tk.Frame(departments, bg=LIGHT); form.pack(fill="x", padx=10, pady=10)
        tk.Label(form, text="Code (blank = automatic)", bg=LIGHT).pack(side="left"); tk.Entry(form, textvariable=self.dep_code, width=10).pack(side="left", padx=(4, 10))
        tk.Label(form, text="Department Name", bg=LIGHT).pack(side="left"); tk.Entry(form, textvariable=self.dep_name, width=30).pack(side="left", padx=(4, 10))
        tk.Checkbutton(form, text="Active", variable=self.dep_active, bg=LIGHT).pack(side="left", padx=4)
        self.action_button(form, "New", self.new_department).pack(side="left", padx=3)
        tk.Button(form, text="Save", command=self.save_department, bg=GOLD, fg=NAVY, border=0, padx=16, pady=7, font=("Segoe UI", 9, "bold")).pack(side="left", padx=3)
        self.departments_tree = self.table(departments, [("code", "Code", 90), ("name", "Department", 300), ("active", "Active", 70)])
        self.departments_tree.bind("<Double-1>", lambda _e: self.edit_department())
        self.proj_id = None
        self.proj_vars = {k: tk.StringVar() for k in ("code", "name", "party", "start_date", "end_date", "notes")}; self.proj_status = tk.StringVar(value="open"); self.proj_active = tk.BooleanVar(value=True)
        form = tk.LabelFrame(projects, text="Project", bg=LIGHT, padx=8, pady=6); form.pack(fill="x", padx=10, pady=10)
        entries = (("code", "Code (blank = automatic)", 12), ("name", "Project Name", 30), ("start_date", "Start Date", 11), ("end_date", "End Date", 11), ("notes", "Notes", 30))
        for index, (key, label, width) in enumerate(entries):
            tk.Label(form, text=label, bg=LIGHT).grid(row=index // 3, column=(index % 3) * 2, sticky="w", padx=4, pady=3)
            widget = self.date_entry(form, self.proj_vars[key], width) if key.endswith("date") else tk.Entry(form, textvariable=self.proj_vars[key], width=width)
            widget.grid(row=index // 3, column=(index % 3) * 2 + 1, sticky="w", padx=4, pady=3)
        tk.Label(form, text="Customer", bg=LIGHT).grid(row=1, column=4, sticky="w", padx=4)
        self.proj_party_box = ttk.Combobox(form, textvariable=self.proj_vars["party"], width=28); self.proj_party_box.grid(row=1, column=5, sticky="w", padx=4)
        tk.Label(form, text="Status", bg=LIGHT).grid(row=2, column=0, sticky="w", padx=4)
        ttk.Combobox(form, textvariable=self.proj_status, values=["open", "on hold", "completed", "cancelled"], state="readonly", width=12).grid(row=2, column=1, sticky="w", padx=4)
        tk.Checkbutton(form, text="Active", variable=self.proj_active, bg=LIGHT).grid(row=2, column=2, sticky="w")
        buttons = tk.Frame(form, bg=LIGHT); buttons.grid(row=2, column=3, columnspan=3, sticky="w")
        self.action_button(buttons, "New", self.new_project).pack(side="left", padx=3)
        tk.Button(buttons, text="Save", command=self.save_project, bg=GOLD, fg=NAVY, border=0, padx=16, pady=7, font=("Segoe UI", 9, "bold")).pack(side="left", padx=3)
        self.projects_tree = self.table(projects, [("code", "Code", 95), ("name", "Project", 230), ("party", "Customer", 180), ("start", "Start", 90), ("end", "End", 90), ("status", "Status", 90), ("active", "Active", 60)])
        self.projects_tree.bind("<Double-1>", lambda _e: self.edit_project())
        self.load_dimensions_pages()

    def load_dimensions_pages(self):
        if not hasattr(self, "departments_tree"): return
        lists = self.dimension_lists(refresh=True)
        self.departments_tree.delete(*self.departments_tree.get_children())
        for d in lists["departments"]: self.departments_tree.insert("", "end", iid=str(d["id"]), values=(d["code"], d["name"], "Yes" if d["active"] else "No"))
        self.projects_tree.delete(*self.projects_tree.get_children())
        for p in lists["projects"]:
            self.projects_tree.insert("", "end", iid=str(p["id"]), values=(p["code"], p["name"], p.get("party_name") or "", self._dd(p.get("start_date")), self._dd(p.get("end_date")), p["status"].title(), "Yes" if p["active"] else "No"))
        try: self.proj_party_box["values"] = [p["name"] for p in self.client.parties() if p["kind"] in ("customer", "both")]
        except Exception: pass

    @staticmethod
    def _dd(value):
        text = str(value or "")
        return f"{text[8:10]}-{text[5:7]}-{text[:4]}" if len(text) == 10 and text[4] == "-" else text

    def new_department(self): self.dep_id = None; self.dep_code.set(""); self.dep_name.set(""); self.dep_active.set(True)

    def edit_department(self):
        selected = self.departments_tree.selection()
        if not selected: return
        d = next(d for d in self.dimension_lists()["departments"] if str(d["id"]) == selected[0])
        self.dep_id = d["id"]; self.dep_code.set(d["code"]); self.dep_name.set(d["name"]); self.dep_active.set(bool(d["active"]))

    def save_department(self):
        try: saved = self.client.save_department({"id": self.dep_id, "code": self.dep_code.get(), "name": self.dep_name.get(), "active": self.dep_active.get()})
        except Exception as exc: return messagebox.showerror("Departments", str(exc))
        self.new_department(); self.load_dimensions_pages(); messagebox.showinfo("Departments", f'Department {saved["code"]} - {saved["name"]} saved')

    def new_project(self):
        self.proj_id = None; [v.set("") for v in self.proj_vars.values()]; self.proj_status.set("open"); self.proj_active.set(True)

    def edit_project(self):
        selected = self.projects_tree.selection()
        if not selected: return
        p = next(p for p in self.dimension_lists()["projects"] if str(p["id"]) == selected[0]); self.proj_id = p["id"]
        for key, value in (("code", p["code"]), ("name", p["name"]), ("party", p.get("party_name") or ""), ("start_date", self._dd(p.get("start_date"))),
                           ("end_date", self._dd(p.get("end_date"))), ("notes", p.get("notes") or "")): self.proj_vars[key].set(value)
        self.proj_status.set(p["status"]); self.proj_active.set(bool(p["active"]))

    def save_project(self):
        payload = {key: var.get().strip() for key, var in self.proj_vars.items()}
        payload.update(id=self.proj_id, party_name=payload.pop("party"), status=self.proj_status.get(), active=self.proj_active.get())
        try: saved = self.client.save_project(payload)
        except Exception as exc: return messagebox.showerror("Projects", str(exc))
        self.new_project(); self.load_dimensions_pages(); messagebox.showinfo("Projects", f'Project {saved["code"]} - {saved["name"]} saved')

    # ------------------------------------------------------------ budget page
    def build_budget_page(self, nested):
        page = tk.Frame(nested, bg=LIGHT); nested.add(page, text="Budget")
        year = getattr(self, "current_fiscal_year", datetime.now().year)
        self.budget_year = tk.StringVar(value=str(year)); self.budget_currency = tk.StringVar(value="USD")
        self.budget_department = tk.StringVar(value=NONE); self.budget_project = tk.StringVar(value=NONE); self.budget_upto = tk.StringVar(value="Dec")
        bar = tk.Frame(page, bg=LIGHT); bar.pack(fill="x", padx=10, pady=8)
        tk.Label(bar, text="Year", bg=LIGHT).pack(side="left"); tk.Entry(bar, textvariable=self.budget_year, width=6).pack(side="left", padx=(4, 10))
        tk.Label(bar, text="Currency", bg=LIGHT).pack(side="left")
        ttk.Combobox(bar, textvariable=self.budget_currency, values=["USD", "LBP", "EUR", "AED"], state="readonly", width=6).pack(side="left", padx=(4, 10))
        self.dimension_selectors(bar, self.budget_department, self.budget_project)
        tk.Button(bar, text="Load", command=self.load_budget, bg=GOLD, fg=NAVY, border=0, padx=14, pady=6, font=("Segoe UI", 9, "bold")).pack(side="left", padx=3)
        tk.Button(bar, text="Save Budget", command=self.save_budget, bg=NAVY, fg="white", border=0, padx=14, pady=6).pack(side="left", padx=3)
        tools = tk.Frame(page, bg=LIGHT); tools.pack(fill="x", padx=10)
        self.action_button(tools, "Add Account", self.add_budget_line).pack(side="left", padx=(0, 3))
        tk.Button(tools, text="Delete Line", command=self.delete_budget_line, bg=RED, fg="white", border=0, padx=12, pady=7).pack(side="left", padx=3)
        self.action_button(tools, "Spread Annual over 12 Months", self.spread_budget_line).pack(side="left", padx=3)
        tk.Label(tools, text="Compare up to", bg=LIGHT).pack(side="left", padx=(20, 4))
        ttk.Combobox(tools, textvariable=self.budget_upto, values=MONTHS, state="readonly", width=5).pack(side="left")
        tk.Button(tools, text="Budget vs Actual", command=self.budget_vs_actual, bg=GOLD, fg=NAVY, border=0, padx=14, pady=6, font=("Segoe UI", 9, "bold")).pack(side="left", padx=6)
        tk.Label(page, text="Type an account number (F2 to search), then either an Annual amount or the monthly amounts. An annual budget alone is spread evenly over the year. "
                 "Leave Department and Project as (none) for the company budget.", bg=LIGHT, fg=MUTED, wraplength=1100, justify="left").pack(fill="x", padx=12, pady=(4, 0))
        columns = [("line", "#", 40, "center"), ("account", "Account", 105, "w"), ("account_name", "Account Name", 190, "w"), ("annual", "Annual", 105, "e")] + \
                  [(m.lower(), m, 78, "e") for m in MONTHS] + [("total", "Months Total", 105, "e")]
        self.budget_sheet = EditableSheet(self, page, columns, ["account", "annual"] + [m.lower() for m in MONTHS], self.budget_cell_changed, height=9, lookup_column="account")
        self.budget_total_label = tk.Label(page, text="", bg=LIGHT, fg=NAVY, font=("Segoe UI", 10, "bold")); self.budget_total_label.pack(anchor="e", padx=14)
        self.budget_viewer = self.report_viewer(page)

    def budget_row(self, code="", name="", annual=0.0, months=None):
        row = {"account": code, "account_name": name, "annual": annual, **{m.lower(): v for m, v in zip(MONTHS, months or [0.0] * 12)}}
        return self.format_budget_row(row)

    def format_budget_row(self, row):
        total = sum(_num(row[m.lower()]) or 0 for m in MONTHS); row["total"] = f"{total:,.2f}"
        row["_display"] = {"annual": f"{_num(row['annual']) or 0:,.2f}", **{m.lower(): (f"{_num(row[m.lower()]):,.2f}" if _num(row[m.lower()]) else "") for m in MONTHS}}
        return row

    def budget_cell_changed(self, iid, key, text):
        row = self.budget_sheet.rows[iid]
        if key == "account":
            code = text.split(" - ", 1)[0].strip(); account = self.account_by_code(code) if code else None
            if code and not account: messagebox.showwarning("Budget", f"Account {code} was not found. Press F2 in the cell to search."); return False
            row["account"] = code; row["account_name"] = account["name_en"] if account else ""
        else:
            value = _num(text)
            if value is None or value < 0: messagebox.showwarning("Budget", "Enter a positive number"); return False
            row[key] = value
        self.format_budget_row(row); self.update_budget_total()

    def update_budget_total(self):
        rows = self.budget_sheet.ordered(); annual = sum(_num(r["annual"]) or 0 for r in rows); months = sum(sum(_num(r[m.lower()]) or 0 for m in MONTHS) for r in rows)
        self.budget_total_label.config(text=f"Total annual: {annual:,.2f}    Total of months: {months:,.2f} {self.budget_currency.get()}")

    def add_budget_line(self):
        iid = self.budget_sheet.insert(self.budget_row()); self.budget_sheet.tree.selection_set(iid); self.budget_sheet.tree.focus(iid)
        self.after(30, lambda: self.budget_sheet.edit(iid, "account"))

    def delete_budget_line(self):
        if not self.budget_sheet.delete_selected(): return messagebox.showwarning("Budget", "Select a line first")
        self.update_budget_total()

    def spread_budget_line(self):
        iid, row = self.budget_sheet.selected()
        if not row: return messagebox.showwarning("Budget", "Select a line first")
        annual = _num(row["annual"]) or 0
        if not annual: return messagebox.showwarning("Budget", "Enter the Annual amount first")
        share = round(annual / 12, 2)
        for index, month in enumerate(MONTHS): row[month.lower()] = share if index < 11 else round(annual - share * 11, 2)
        self.format_budget_row(row); self.budget_sheet.refresh(iid); self.update_budget_total()

    def budget_filters(self):
        try: year = int(self.budget_year.get())
        except ValueError: raise ValueError("Enter the budget year, for example 2025")
        return year, self.budget_currency.get(), self.dimension_code(self.budget_department.get()), self.dimension_code(self.budget_project.get())

    def load_budget(self):
        try: year, currency, department, project = self.budget_filters(); lines = self.client.budgets(year, currency, department or None, project or None)
        except Exception as exc: return messagebox.showerror("Budget", str(exc))
        self.budget_sheet.clear()
        for line in lines: self.budget_sheet.insert(self.budget_row(line["account_code"], line["account_name"], line["annual"], line["months"]))
        if not lines: self.budget_sheet.insert(self.budget_row())
        self.update_budget_total()

    def save_budget(self):
        try: year, currency, department, project = self.budget_filters()
        except Exception as exc: return messagebox.showerror("Budget", str(exc))
        lines = [{"account_code": r["account"], "annual": _num(r["annual"]) or 0, "months": [_num(r[m.lower()]) or 0 for m in MONTHS]} for r in self.budget_sheet.ordered() if r["account"]]
        for r in lines:
            if r["annual"] and any(r["months"]) and abs(sum(r["months"]) - r["annual"]) > 0.01:
                if not messagebox.askyesno("Budget", f"Account {r['account_code']}: the months add up to {sum(r['months']):,.2f}, not the annual {r['annual']:,.2f}. The monthly amounts will be used. Continue?"): return
        try: self.client.save_budget({"year": year, "currency": currency, "department": department, "project": project, "lines": lines})
        except Exception as exc: return messagebox.showerror("Budget", str(exc))
        self.load_budget(); messagebox.showinfo("Budget", f"Budget {year} ({currency}) saved for {len(lines)} account(s)")

    def budget_vs_actual(self):
        try: year, currency, department, project = self.budget_filters()
        except Exception as exc: return messagebox.showerror("Budget", str(exc))
        month = MONTHS.index(self.budget_upto.get()) + 1
        import calendar
        options = {"date_from": f"01-01-{year}", "date_to": f"{calendar.monthrange(year, month)[1]:02d}-{month:02d}-{year}", "first_column": currency, "second_column": "none",
                   "budget": True, "carry_forward": False, "profit_loss_only": True, "department": department, "project": project, "posting_status": "posted"}
        try: result = self.client.account_report(options)
        except Exception as exc: return messagebox.showerror("Budget", str(exc))
        self.budget_result = result; self.show_sections(self.budget_viewer, result["sections"])

    # ------------------------------------------------------------ balance panel additions
    def add_dimension_options(self, state, parent):
        box = tk.Frame(parent, bg=LIGHT); box.pack(fill="x", pady=(4, 0))
        state["vars"]["department"] = tk.StringVar(value="All"); state["vars"]["project"] = tk.StringVar(value="All")
        self.dimension_selectors(box, state["vars"]["department"], state["vars"]["project"], include_all=True)
        for name, label in (("with_department", "Show Department / Project"), ("split_by_department", "Split by Department"),
                            ("split_by_project", "Split by Project"), ("budget", "Compare with Budget")):
            state["flags"][name] = tk.BooleanVar(value=False); tk.Checkbutton(box, text=label, variable=state["flags"][name], bg=LIGHT).pack(side="left", padx=4)
