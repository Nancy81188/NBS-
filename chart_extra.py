"""Accounts the company asked for (created automatically when missing) and the default posting accounts."""

EXTRA_ACCOUNTS = [
    ("44210", "VAT on Purchases - Deductible", "asset"), ("44211", "VAT on Export-related Purchases - Recoverable", "asset"),
    ("44216", "VAT on Expenses - Deductible", "asset"),
    ("6313", "Commissions", "expense"), ("6315", "Schooling Allowances", "expense"), ("6319", "Transport Allowances", "expense"),
    ("601800001", "Purchase Costs - Freight", "expense"), ("601800002", "Purchase Costs - Insurance", "expense"),
    ("601800003", "Purchase Costs - Customs Duties", "expense"), ("601800004", "Purchase Costs - Customs Broker Fees", "expense"),
    ("601800005", "Purchase Costs - Other", "expense"),
]
SALES_VAT = "4427"; PURCHASE_VAT = "44210"; EXPENSE_VAT = "44216"; EXPORT_VAT = "44211"
LANDED_COST_ACCOUNTS = {"freight": "601800001", "insurance": "601800002", "customs_duties": "601800003", "broker_fees": "601800004", "other_costs": "601800005"}
PAYROLL_MAP = {"salary": "6311", "overtime": "6311", "retro_salary": "6311", "bonus": "6312", "thirteenth_month": "6312", "commission": "6313",
               "schooling": "6315", "transport": "6319", "tax": "4411", "nssf": "4431", "payable": "421100001"}
MANAGER_PAYROLL_MAP = {**PAYROLL_MAP, "salary": "6316", "overtime": "6316", "retro_salary": "6316"}
OLD_PAYROLL_MAP = {"salary": "621100001", "transport": "621100003", "overtime": "621100004", "commission": "621100005", "retro_salary": "621100006",
                   "schooling": "621100007", "bonus": "621100008", "thirteenth_month": "621100009", "tax": "443100001", "nssf": "447100001", "payable": "421100001"}


def ensure_accounts(db):
    for code, name, kind in EXTRA_ACCOUNTS:
        if db.execute("SELECT 1 FROM accounts WHERE code=?", (code,)).fetchone(): continue
        parent = None
        for size in range(len(code) - 1, 0, -1):
            parent = db.execute("SELECT id FROM accounts WHERE code=?", (code[:size],)).fetchone()
            if parent: break
        db.execute("INSERT INTO accounts(code,name_en,type,parent_id) VALUES(?,?,?,?)", (code, name, kind, parent["id"] if parent else None))
