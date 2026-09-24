"""Lebanese payroll official reports (R5, R6, R10) built from saved payroll records.

Every figure is taken from the payroll records as they were calculated, so each month keeps
the tax brackets, NSSF rates and ceilings that were effective on its own period date.
Official amounts are shown in LBP; records kept in another currency are converted with the
exchange rate of their payroll period.
"""
from __future__ import annotations

import calendar
from datetime import date
from decimal import Decimal, ROUND_HALF_UP

REPORTS = {
    "R10": "R10 - Quarterly Salary Tax Withholding Return",
    "R5": "R5 - Annual Salary Tax Declaration (Employer Summary)",
    "R6": "R6 - Individual Annual Salary Statement",
}
PERIOD_TYPES = ("monthly", "quarterly", "yearly")
GROUPS = {"employee": "Employees", "manager": "Managers"}
COMPONENTS = (("salary", "Salary"), ("transport", "Transport"), ("overtime", "Overtime"), ("commission", "Commission"),
              ("retro_salary", "Retro Salary"), ("schooling", "Schooling"), ("bonus", "Bonus"), ("thirteenth_month", "13th Salary"))
MONEY_FIELDS = tuple(key for key, _ in COMPONENTS) + ("gross_salary", "taxable_salary", "nssf_base", "employee_nssf",
    "employer_medical", "employer_family", "employer_end_service", "net_salary", "retro_tax")
ZERO = Decimal("0")


def period_range(period_type, year, index=1):
    period_type = str(period_type or "").lower()
    if period_type not in PERIOD_TYPES: raise ValueError("Period must be Monthly, Quarterly or Yearly")
    try: year = int(year); index = int(index or 1)
    except (TypeError, ValueError) as exc: raise ValueError("Enter a valid year and period") from exc
    if year < 2000 or year > 2100: raise ValueError("Enter a valid year")
    if period_type == "monthly":
        if not 1 <= index <= 12: raise ValueError("Month must be between 1 and 12")
        start, end = date(year, index, 1), date(year, index, calendar.monthrange(year, index)[1])
        label = f"{calendar.month_name[index]} {year}"
    elif period_type == "quarterly":
        if not 1 <= index <= 4: raise ValueError("Quarter must be between 1 and 4")
        first = 3 * index - 2
        start, end = date(year, first, 1), date(year, first + 2, calendar.monthrange(year, first + 2)[1])
        label = f"Q{index} {year} ({calendar.month_abbr[first]} - {calendar.month_abbr[first + 2]})"
    else:
        start, end = date(year, 1, 1), date(year, 12, 31); label = f"Year {year}"
    return start.isoformat(), end.isoformat(), label


def _lbp(value):
    return Decimal(str(value or 0)).quantize(Decimal("1"), rounding=ROUND_HALF_UP)


def _rate_text(value):
    rate = Decimal(str(value or 0)) * 100
    return f"{rate.normalize():f}%"


def _ceiling_text(value):
    amount = Decimal(str(value or 0))
    return "No ceiling" if amount <= 0 else amount.quantize(Decimal("1"))


def _load_records(db, start, end, include_drafts):
    status = "" if include_drafts else " AND p.status='posted'"
    with db.connect() as connection:
        rows = [dict(row) for row in connection.execute(f"""SELECT p.*,e.employee_number,e.full_name,e.mof_number,e.nssf_number,
            e.national_id,e.marital_status,e.spouse_works,e.children,e.employee_group,e.job_title,e.hire_date,e.leave_date
            FROM payroll_records p JOIN employees e ON e.id=p.employee_id
            WHERE p.period_date>=? AND p.period_date<=?{status}
            ORDER BY e.employee_number,p.period_date""", (start, end))]
    cache = {}
    for row in rows:
        key = (row["currency"], row["period_date"])
        if key not in cache: cache[key] = db._converted_amount(Decimal("1"), row["currency"], "LBP", row["period_date"])
        rate = cache[key]; row["lbp_rate"] = rate
        lbp = {field: _lbp(Decimal(str(row.get(field) or 0)) * rate) for field in MONEY_FIELDS}
        stored_tax_lbp = Decimal(str(row.get("income_tax_lbp") or 0))
        lbp["income_tax"] = _lbp(stored_tax_lbp if stored_tax_lbp or row["currency"] == "LBP" else Decimal(str(row.get("income_tax") or 0)) * rate)
        lbp["family_deductions"] = max(ZERO, lbp["gross_salary"] - lbp["taxable_salary"])
        lbp["employer_total"] = lbp["employer_medical"] + lbp["employer_family"] + lbp["employer_end_service"]
        lbp["nssf_total"] = lbp["employer_total"] + lbp["employee_nssf"]
        row["lbp"] = lbp
    return rows


def _group_keys(group):
    group = str(group or "both").lower()
    if group in ("both", "all", "separate"): return ["employee", "manager"]
    if group not in GROUPS: raise ValueError("Group must be Employees, Managers or Both")
    return [group]


def _employee_totals(records):
    employees = {}
    for row in records:
        item = employees.setdefault(row["employee_id"], {"row": row, "months": set(), "totals": {}})
        item["months"].add(row["period_date"][:7])
        for key, value in row["lbp"].items(): item["totals"][key] = item["totals"].get(key, ZERO) + value
    return list(employees.values())


def _sum(items, key):
    return sum((item["totals"].get(key, ZERO) for item in items), ZERO)


def _family_text(row):
    status = "Married" if str(row.get("marital_status") or "").lower() in ("married", "spouse") else "Single"
    spouse = ", spouse works" if status == "Married" and int(row.get("spouse_works") or 0) else ""
    return f"{status}{spouse}, {int(row.get('children') or 0)} child(ren)"


def _settings_section(db, start, end):
    with db.connect() as connection:
        rows = [dict(row) for row in connection.execute("""SELECT * FROM payroll_settings
            WHERE date_from<=? AND (date_to IS NULL OR date_to='' OR date_to>=?) ORDER BY date_from""", (end, start))]
    headers = ["Effective From", "Effective To", "Employee Rate", "Employee Ceiling", "Medical Rate", "Medical Ceiling",
               "Family Rate", "Family Ceiling", "End-of-Service Rate", "End-of-Service Ceiling"]
    body = []
    for row in rows:
        body.append([_display(row["date_from"]), _display(row.get("date_to")) or "Open", _rate_text(row["employee_nssf_rate"]), _ceiling_text(row["employee_ceiling"]),
            _rate_text(row["medical_rate"]), _ceiling_text(row["medical_ceiling"]), _rate_text(row["family_rate"]), _ceiling_text(row["family_ceiling"]),
            _rate_text(row["end_service_rate"]), _ceiling_text(row["end_service_ceiling"])])
    return {"heading": "NSSF rates and ceilings applied (monthly ceilings in LBP, by effective date)", "headers": headers, "rows": body, "total_rows": []}, rows


def _display(value):
    text = str(value or "")
    return f"{text[8:10]}-{text[5:7]}-{text[:4]}" if len(text) == 10 and text[4] == "-" else text


def _employee_rate_label(settings_rows):
    rates = sorted({_rate_text(row["employee_nssf_rate"]) for row in settings_rows})
    return f"Employee NSSF ({' / '.join(rates)})" if rates else "Employee NSSF"


def _nssf_section(label, items, employee_label):
    headers = ["Emp. No.", "Employee Name", "NSSF No.", "Salary Subject to NSSF", employee_label, "Employer Medical",
               "Employer Family", "Employer End-of-Service", "Total Employer", "Total NSSF"]
    fields = ("nssf_base", "employee_nssf", "employer_medical", "employer_family", "employer_end_service", "employer_total", "nssf_total")
    rows = [[item["row"]["employee_number"], item["row"]["full_name"], item["row"].get("nssf_number") or ""] + [item["totals"][f] for f in fields] for item in items]
    rows.append(["TOTAL", f"{len(items)} {label.lower()}", ""] + [_sum(items, f) for f in fields])
    return {"heading": f"NSSF contributions - {label}", "headers": headers, "rows": rows, "total_rows": [len(rows) - 1]}


def _r10_sections(label, items):
    headers = ["Emp. No.", "Employee Name", "MOF No.", "Months"] + [name for _, name in COMPONENTS] + ["Gross", "Family Deductions", "Taxable", "Income Tax", "of which Retro Tax"]
    fields = [key for key, _ in COMPONENTS] + ["gross_salary", "family_deductions", "taxable_salary", "income_tax", "retro_tax"]
    rows = [[item["row"]["employee_number"], item["row"]["full_name"], item["row"].get("mof_number") or "", len(item["months"])] + [item["totals"][f] for f in fields] for item in items]
    rows.append(["TOTAL", f"{len(items)} {label.lower()}", "", ""] + [_sum(items, f) for f in fields])
    return [{"heading": f"Salary tax withheld - {label}", "headers": headers, "rows": rows, "total_rows": [len(rows) - 1]}]


def _r5_sections(label, items):
    summary = [[name, _sum(items, key)] for key, name in COMPONENTS]
    summary += [["Total gross salaries and benefits", _sum(items, "gross_salary")], ["Family deductions", _sum(items, "family_deductions")],
        ["Total taxable income", _sum(items, "taxable_salary")], ["Salary tax withheld", _sum(items, "income_tax")],
        ["of which tax on retroactive salary", _sum(items, "retro_tax")], ["Employee NSSF withheld", _sum(items, "employee_nssf")],
        ["Employer NSSF contributions", _sum(items, "employer_total")], ["Number of employees", len(items)]]
    headers = ["Emp. No.", "Employee Name", "MOF No.", "NSSF No.", "Family Status", "Months", "Gross", "Family Deductions", "Taxable", "Income Tax"]
    fields = ("gross_salary", "family_deductions", "taxable_salary", "income_tax")
    rows = [[item["row"]["employee_number"], item["row"]["full_name"], item["row"].get("mof_number") or "", item["row"].get("nssf_number") or "",
             _family_text(item["row"]), len(item["months"])] + [item["totals"][f] for f in fields] for item in items]
    rows.append(["TOTAL", f"{len(items)} {label.lower()}", "", "", "", ""] + [_sum(items, f) for f in fields])
    return [{"heading": f"Declaration summary - {label}", "headers": ["Item", "Amount (LBP)"], "rows": summary, "total_rows": [8, 10, 11]},
            {"heading": f"Employee schedule - {label}", "headers": headers, "rows": rows, "total_rows": [len(rows) - 1]}]


def _r6_sections(label, records, employee_label):
    sections = []
    headers = ["Month"] + [name for _, name in COMPONENTS] + ["Retro Period", "Gross", "Taxable", "Income Tax", "Retro Tax", employee_label, "Net Salary", "Status"]
    fields = [key for key, _ in COMPONENTS]
    tail = ("gross_salary", "taxable_salary", "income_tax", "retro_tax", "employee_nssf", "net_salary")
    by_employee = {}
    for row in records: by_employee.setdefault(row["employee_id"], []).append(row)
    for rows in by_employee.values():
        first = rows[0]; body = []
        for row in rows:
            retro = f"{_display(row.get('retro_from'))} to {_display(row.get('retro_to'))}" if row.get("retro_from") else ""
            body.append([row["period_date"][5:7] + "-" + row["period_date"][:4]] + [row["lbp"][f] for f in fields] + [retro] + [row["lbp"][f] for f in tail] + [row["status"]])
        totals = ["TOTAL"] + [sum((row["lbp"][f] for row in rows), ZERO) for f in fields] + [""] + [sum((row["lbp"][f] for row in rows), ZERO) for f in tail] + [""]
        body.append(totals)
        heading = (f"{label[:-1] if label.endswith('s') else label} {first['employee_number']} - {first['full_name']} | MOF {first.get('mof_number') or '-'} | "
                   f"NSSF {first.get('nssf_number') or '-'} | {_family_text(first)}")
        sections.append({"heading": heading, "headers": headers, "rows": body, "total_rows": [len(body) - 1]})
    return sections


def build_payroll_report(db, report="R10", period_type="quarterly", year=None, index=1, group="both", include_drafts=False):
    report = str(report or "R10").upper()
    if report not in REPORTS: raise ValueError("Report must be R5, R6 or R10")
    start, end, label = period_range(period_type, year or date.today().year, index)
    records = _load_records(db, start, end, bool(include_drafts))
    settings_section, settings_rows = _settings_section(db, start, end)
    employee_label = _employee_rate_label(settings_rows)
    company = db.settings()
    sections = []; summary = {}
    for key in _group_keys(group):
        group_label = GROUPS[key]
        group_records = [row for row in records if (row.get("employee_group") or "employee") == key]
        items = _employee_totals(group_records)
        summary[key] = {"employees": len(items), "gross": float(_sum(items, "gross_salary")), "income_tax": float(_sum(items, "income_tax")),
                        "retro_tax": float(_sum(items, "retro_tax")), "employee_nssf": float(_sum(items, "employee_nssf")), "employer_nssf": float(_sum(items, "employer_total"))}
        if not items:
            sections.append({"heading": f"{group_label}", "headers": ["Note"], "rows": [[f"No {'posted ' if not include_drafts else ''}payroll for {group_label.lower()} in {label}"]], "total_rows": []})
            continue
        if report == "R10": sections += _r10_sections(group_label, items)
        elif report == "R5": sections += _r5_sections(group_label, items)
        else: sections += _r6_sections(group_label, group_records, employee_label)
        sections.append(_nssf_section(group_label, items, employee_label))
    if len(summary) > 1:
        combined = [["Employees and managers", sum(v["employees"] for v in summary.values()), sum(v["gross"] for v in summary.values()),
                     sum(v["income_tax"] for v in summary.values()), sum(v["retro_tax"] for v in summary.values()),
                     sum(v["employee_nssf"] for v in summary.values()), sum(v["employer_nssf"] for v in summary.values())]]
        rows = [[GROUPS[k], v["employees"], v["gross"], v["income_tax"], v["retro_tax"], v["employee_nssf"], v["employer_nssf"]] for k, v in summary.items()] + combined
        sections.append({"heading": "Grand total", "headers": ["Group", "Employees", "Gross", "Income Tax", "Retro Tax", employee_label, "Employer NSSF"],
                         "rows": [[r[0], r[1]] + [_lbp(x) for x in r[2:]] for r in rows], "total_rows": [len(rows) - 1]})
    sections.append(settings_section)
    meta = [f"Company: {company.get('company_name') or '-'}   MOF No.: {company.get('company_mof') or '-'}",
            f"Period: {label} ({_display(start)} to {_display(end)})   Amounts in LBP",
            "Source: " + ("posted and draft payroll (draft figures are not final)" if include_drafts else "posted payroll only")]
    return {"report": report, "title": REPORTS[report], "period_label": label, "date_from": start, "date_to": end,
            "meta": meta, "sections": sections, "summary": summary, "record_count": len(records)}


def json_ready(result):
    """Convert Decimals for the JSON API."""
    def convert(value):
        if isinstance(value, Decimal): return float(value)
        if isinstance(value, list): return [convert(v) for v in value]
        if isinstance(value, dict): return {k: convert(v) for k, v in value.items()}
        return value
    return convert(result)
