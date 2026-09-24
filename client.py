from __future__ import annotations

import json
import base64
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

class SessionExpired(RuntimeError):
    """Raised when the server no longer accepts the saved sign-in."""


class ApiClient:
    def __init__(self, base_url="http://127.0.0.1:8765", on_unauthorized=None):
        self.base_url = base_url.rstrip("/")
        self.token = None
        self.company_id = None
        self.fiscal_year = None
        self.on_unauthorized = on_unauthorized

    def request(self, method, path, body=None):
        data = json.dumps(body).encode("utf-8") if body is not None else None
        headers = {"Content-Type": "application/json"}
        if self.token:
            headers["Authorization"] = f"Bearer {self.token}"
        if self.company_id: headers["X-Company-ID"]=str(self.company_id)
        if self.fiscal_year: headers["X-Fiscal-Year"]=str(self.fiscal_year)
        request = Request(self.base_url + path, data=data, headers=headers, method=method)
        try:
            with urlopen(request, timeout=60) as response:
                return json.loads(response.read().decode("utf-8"))
        except HTTPError as exc:
            try: message = json.loads(exc.read().decode("utf-8")).get("error", str(exc))
            except Exception: message = f"The server returned an error ({exc.code}). Please try again."
            if exc.code == 401 and path != "/api/login" and self.token:
                self.token = None
                if self.on_unauthorized:
                    try: self.on_unauthorized()
                    except Exception: pass
                raise SessionExpired("Your session has ended. Please sign in again.") from exc
            raise RuntimeError(message) from exc
        except URLError as exc:
            raise RuntimeError(f"Cannot reach the Saber Accounting data service at {self.base_url}. "
                               "Check that the server computer is on and the address is correct.") from exc
        except TimeoutError as exc:
            raise RuntimeError("The data service took too long to answer. Please try again.") from exc

    def login(self, username, password):
        result = self.request("POST", "/api/login", {"username": username, "password": password})
        self.token = result["token"]
        return result

    def dashboard(self): return self.request("GET", "/api/dashboard")["items"]
    def companies(self): return self.request("GET","/api/companies")["items"]
    def create_company(self,item): return self.request("POST","/api/companies",item)["company"]
    def update_company(self,company_id,item): return self.request("PUT",f"/api/companies/{company_id}",item)["company"]
    def create_fiscal_year(self,company_id,year): return self.request("POST","/api/companies/year",{"company_id":company_id,"year":year})["company"]
    def select_company_year(self,company_id,year): self.company_id=company_id; self.fiscal_year=int(year)
    def professional_dashboard(self): return self.request("GET","/api/dashboard/professional")
    def invoices(self): return self.request("GET", "/api/invoices")["items"]
    def accounts(self): return self.request("GET", "/api/accounts")["items"]
    def next_account_number(self,prefix): return self.request("GET",f"/api/accounts/next-number?{urlencode({'prefix':prefix})}")["account_number"]
    def save_account(self,item): return self.request("POST","/api/accounts",item)["account"]
    def rename_account(self,code,name): return self.request("PUT",f"/api/accounts/{code}",{"name_en":name})["account"]
    def parties(self): return self.request("GET", "/api/parties")["items"]
    def branches(self): return self.request("GET","/api/branches")["items"]
    def save_branch(self,name): return self.request("POST","/api/branches",{"name":name})["branch"]
    def statement(self, party_id, from_date=None, to_date=None, currency=None, include_opening=True, display_currency=None, branch_id=None):
        query = urlencode({key:value for key,value in {
            "party_id":party_id,"from_date":from_date,"to_date":to_date,"currency":currency,
            "include_opening":"true" if include_opening else "false","display_currency":display_currency,"branch_id":branch_id
        }.items() if value})
        return self.request("GET", f"/api/statement?{query}")
    def add_invoice_item(self, invoice_id, item):
        return self.request("POST", f"/api/invoices/{invoice_id}/items", {"item":item})["invoice"]
    def trial_balance(self, from_date=None, to_date=None, account=None, include_subaccounts=True, account_from=None, account_to=None, branch_id=None, posting_status="posted"):
        query = urlencode({key: value for key, value in {
            "from_date": from_date, "to_date": to_date,"account":account,"include_subaccounts":"true" if include_subaccounts else "false","account_from":account_from,"account_to":account_to,"branch_id":branch_id,"posting_status":posting_status
        }.items() if value})
        path = "/api/trial-balance" + (f"?{query}" if query else "")
        return self.request("GET", path)["items"]
    def journal(self, from_date=None, to_date=None, currency=None):
        query = urlencode({key: value for key, value in {
            "from_date": from_date, "to_date": to_date, "currency": currency
        }.items() if value})
        path = "/api/journal" + (f"?{query}" if query else "")
        return self.request("GET", path)["items"]
    def import_invoices(self, items, replace_existing=True): return self.request("POST", "/api/invoices/import", {"items": items, "replace_existing": replace_existing})
    def create_manual_invoice(self, invoice, items): return self.request("POST", "/api/invoices/manual", {"invoice": invoice, "items": items})
    def update_invoice(self, invoice_id, invoice): return self.request("PUT", f"/api/invoices/{invoice_id}", {"invoice": invoice})
    def delete_invoice(self,invoice_id): return self.request("DELETE",f"/api/invoices/{invoice_id}")
    def delete_journal_voucher(self,entry_id): return self.request("DELETE",f"/api/journal/{entry_id}")
    def delete_opening_voucher(self,entry_id): return self.request("DELETE",f"/api/opening-vouchers/{entry_id}")
    def journal_voucher(self,entry_id): return self.request("GET",f"/api/journal-vouchers/{entry_id}")
    def save_journal_voucher(self,voucher,lines,entry_id=None): return self.request("PUT" if entry_id else "POST",f"/api/journal-vouchers/{entry_id}" if entry_id else "/api/journal-vouchers",{"voucher":voucher,"lines":lines})
    def invoice_detail(self, invoice_id): return self.request("GET",f"/api/invoices/{invoice_id}/detail")
    def cancel_invoice(self, invoice_id, reason): return self.request("POST",f"/api/invoices/{invoice_id}/cancel",{"reason":reason})["invoice"]
    def duplicate_invoice(self, invoice_id): return self.request("POST",f"/api/invoices/{invoice_id}/duplicate",{})["invoice"]
    def invoice_history(self, invoice_id): return self.request("GET",f"/api/invoices/{invoice_id}/history")["items"]
    def attachments(self, invoice_id): return self.request("GET",f"/api/invoices/{invoice_id}/attachments")["items"]
    def upload_attachment(self, invoice_id, file_name, mime_type, content):
        return self.request("POST",f"/api/invoices/{invoice_id}/attachments",{
            "file_name":file_name,"mime_type":mime_type,"content":base64.b64encode(content).decode("ascii")})
    def download_attachment(self, attachment_id):
        result=self.request("GET",f"/api/attachments/{attachment_id}"); result["content"]=base64.b64decode(result["content"]); return result
    def document_cases(self): return self.request("GET","/api/document-cases")["items"]
    def save_document_case(self,item): return self.request("POST","/api/document-cases",item)["case"]
    def post_document_case(self,case_id): return self.request("POST",f"/api/document-cases/{case_id}/post",{})["case"]
    def case_attachments(self,case_id): return self.request("GET",f"/api/document-cases/{case_id}/attachments")["items"]
    def upload_case_attachment(self,case_id,role,file_name,mime_type,content):
        return self.request("POST",f"/api/document-cases/{case_id}/attachments",{"document_role":role,"file_name":file_name,"mime_type":mime_type,"content":base64.b64encode(content).decode("ascii")})
    def download_case_attachment(self,attachment_id):
        result=self.request("GET",f"/api/case-attachments/{attachment_id}"); result["content"]=base64.b64decode(result["content"]); return result
    def party_documents(self,party_id): return self.request("GET",f"/api/parties/{party_id}/documents")["items"]
    def upload_party_document(self,party_id,item,content):
        return self.request("POST",f"/api/parties/{party_id}/documents",{**item,"content":base64.b64encode(content).decode("ascii")})
    def download_party_document(self,document_id):
        result=self.request("GET",f"/api/party-documents/{document_id}"); result["content"]=base64.b64decode(result["content"]); return result
    def profit_loss(self, from_date=None, to_date=None, currency=None):
        query=urlencode({k:v for k,v in {"from_date":from_date,"to_date":to_date,"currency":currency}.items() if v})
        return self.request("GET","/api/profit-loss"+(f"?{query}" if query else ""))["items"]
    def fiscal_years(self): return self.request("GET","/api/fiscal-years")["items"]
    def close_fiscal_year(self, year): return self.request("POST","/api/fiscal-years/close",{"year":year})
    def reopen_fiscal_year(self, year): return self.request("POST","/api/fiscal-years/reopen",{"year":year})
    def refresh_opening(self, source_year): return self.request("POST","/api/fiscal-years/refresh-opening",{"source_year":source_year})
    def fiscal_year_journal(self,year,from_date=None,to_date=None,currency=None):
        query=urlencode({k:v for k,v in {"year":year,"from_date":from_date,"to_date":to_date,"currency":currency}.items() if v})
        return self.request("GET",f"/api/fiscal-year/journal?{query}")["items"]
    def save_party(self, item): return self.request("POST","/api/parties",item)["party"]
    def payments(self): return self.request("GET","/api/payments")["items"]
    def add_payment(self, item): return self.request("POST","/api/payments",item)
    def expenses(self): return self.request("GET","/api/expenses")["items"]
    def add_expense(self, item): return self.request("POST","/api/expenses",item)
    def general_ledger(self, account=None, from_date=None, to_date=None, currency=None):
        query=urlencode({k:v for k,v in {"account":account,"from_date":from_date,"to_date":to_date,"currency":currency}.items() if v})
        return self.request("GET","/api/general-ledger"+(f"?{query}" if query else ""))
    def balance_sheet(self, to_date=None, currency=None):
        query=urlencode({k:v for k,v in {"to_date":to_date,"currency":currency}.items() if v})
        return self.request("GET","/api/balance-sheet"+(f"?{query}" if query else ""))["items"]
    def vat_report(self, from_date=None, to_date=None, currency=None):
        query=urlencode({k:v for k,v in {"from_date":from_date,"to_date":to_date,"currency":currency}.items() if v})
        return self.request("GET","/api/vat-report"+(f"?{query}" if query else ""))
    def cash_flow(self,from_date=None,to_date=None,currency=None):
        query=urlencode({k:v for k,v in {"from_date":from_date,"to_date":to_date,"currency":currency}.items() if v}); return self.request("GET","/api/cash-flow"+(f"?{query}" if query else ""))["items"]
    def aging(self,as_of_date=None,kind=None,currency=None):
        query=urlencode({k:v for k,v in {"as_of_date":as_of_date,"kind":kind,"currency":currency}.items() if v}); return self.request("GET","/api/aging"+(f"?{query}" if query else ""))["items"]
    def comparative_reports(self,from_date,to_date,currency=None):
        query=urlencode({k:v for k,v in {"from_date":from_date,"to_date":to_date,"currency":currency}.items() if v}); return self.request("GET",f"/api/comparative-reports?{query}")
    def users(self): return self.request("GET","/api/users")["items"]
    def save_user(self,item): return self.request("POST","/api/users",item)["user"]
    def backups(self): return self.request("GET","/api/backups")["items"]
    def create_backup(self): return self.request("POST","/api/backups/create",{})
    def restore_backup(self,name): return self.request("POST","/api/backups/restore",{"name":name})
    def settings(self): return self.request("GET","/api/settings")
    def save_settings(self,item): return self.request("POST","/api/settings",item)
    def exchange_rates(self): return self.request("GET","/api/exchange-rates")["items"]
    def save_exchange_rate(self,item): return self.request("POST","/api/exchange-rates",item)
    def restore_euro_rates(self): return self.request("POST","/api/exchange-rates/restore-euro",{})
    def employees(self): return self.request("GET","/api/employees")["items"]
    def next_employee_number(self,prefix="1000"):
        return self.request("GET",f"/api/employees/next-number?{urlencode({'prefix':prefix})}")["employee_number"]
    def save_employee(self,item): return self.request("POST","/api/employees",item)["employee"]
    def payroll(self,from_date=None,to_date=None):
        query=urlencode({k:v for k,v in {"from_date":from_date,"to_date":to_date}.items() if v})
        return self.request("GET","/api/payroll"+(f"?{query}" if query else ""))["items"]
    def calculate_payroll(self,item): return self.request("POST","/api/payroll/calculate",item)
    def save_payroll(self,item): return self.request("POST","/api/payroll",item)["payroll"]
    def post_payroll(self,payroll_id): return self.request("POST",f"/api/payroll/{payroll_id}/post",{})["payroll"]
    def payroll_settings(self,date=None):
        return self.request("GET","/api/payroll/settings"+(f"?{urlencode({'date':date})}" if date else ""))
    def save_payroll_settings(self,item): return self.request("POST","/api/payroll/settings",item)
    def me(self): return self.request("GET","/api/me")
    def payroll_settings_list(self): return self.request("GET","/api/payroll/settings/list")["items"]
    def payroll_report(self,report,period_type,year,index=1,group="both",include_drafts=False):
        query=urlencode({"report":report,"period_type":period_type,"year":year,"index":index,"group":group,"include_drafts":"true" if include_drafts else "false"})
        return self.request("GET",f"/api/payroll/reports?{query}")
    def vat_return(self,year,quarter,currency=None,include_review=False,credit_brought_forward=None):
        query=urlencode({k:v for k,v in {"year":year,"quarter":quarter,"currency":currency,"include_review":"true" if include_review else "false",
            "credit_brought_forward":credit_brought_forward}.items() if v not in (None,"")})
        return self.request("GET",f"/api/vat-return?{query}")
    def vat_returns(self): return self.request("GET","/api/vat-returns")["items"]
    def add_vat_adjustment(self,item): return self.request("POST","/api/vat-return/adjustments",item)
    def delete_vat_adjustment(self,adjustment_id): return self.request("DELETE",f"/api/vat-return/adjustments/{adjustment_id}")
    def save_vat_return(self,year,quarter,credit_brought_forward=None):
        return self.request("POST","/api/vat-return/save",{"year":year,"quarter":quarter,"credit_brought_forward":credit_brought_forward})
    def reopen_vat_return(self,year,quarter): return self.request("POST","/api/vat-return/reopen",{"year":year,"quarter":quarter})
    def set_vat_recoverable(self,source,document_id,recoverable):
        return self.request("POST","/api/vat-recoverable",{"source":source,"id":document_id,"recoverable":bool(recoverable)})
    def document_alerts(self,days=30): return self.request("GET",f"/api/alerts/documents?{urlencode({'days':days})}")
    def next_invoice_number(self,kind="sale",date=None):
        return self.request("GET","/api/invoices/next-number?"+urlencode({k:v for k,v in {"kind":kind,"date":date}.items() if v}))["invoice_number"]
    def next_party_account_number(self,prefix): return self.request("GET",f"/api/parties/next-number?{urlencode({'prefix':prefix})}")["account_number"]
    def invoice_items(self,invoice_id): return self.request("GET",f"/api/invoices/{invoice_id}/items")["items"]
    def account_report(self,options): return self.request("GET","/api/reports/accounts?"+urlencode({"options":json.dumps(options)}))
    def suggested_rates(self,currency,date=None): return self.request("GET","/api/rates/suggest?"+urlencode({k:v for k,v in {"currency":currency,"date":date}.items() if v}))
