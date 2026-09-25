"""Regression checks for data safety and self-service backups."""
import socket
import tempfile
import threading
import time
import unittest
from pathlib import Path

from backup_service import backup_all
from client import ApiClient
from database import Database
from server import run_server
from pdf_import import read_invoice_pdf_pages


class DataSafetyTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.folder = tempfile.TemporaryDirectory()
        cls.database = Path(cls.folder.name) / "saber_accounting_v0_7.db"
        with socket.socket() as probe:
            probe.bind(("127.0.0.1", 0)); cls.port = probe.getsockname()[1]
        cls.url = f"http://127.0.0.1:{cls.port}"
        threading.Thread(target=run_server, kwargs={"host":"127.0.0.1","port":cls.port,"database":str(cls.database),"admin_password":"secret12345"}, daemon=True).start()
        for _ in range(100):
            try: ApiClient(cls.url).login("admin", "secret12345"); break
            except Exception: time.sleep(.05)

    @classmethod
    def tearDownClass(cls): cls.folder.cleanup()

    def test_replacement_failure_preserves_invoice_and_journal(self):
        admin=ApiClient(self.url); admin.login("admin","secret12345")
        company=admin.companies()[0]; admin.select_company_year(company["id"],2024)
        admin.create_manual_invoice({"invoice_date":"01-03-2024","party_name":"Client","kind":"sale","currency":"USD"},
                                    [{"description":"Service","quantity":1,"unit_price":100}])
        voucher=admin.save_journal_voucher({"entry_date":"01-03-2024","description":"Keep me","currency":"USD"},
            [{"account_code":"511","debit":20,"credit":0},{"account_code":"512","debit":0,"credit":20}])
        before=len(admin.invoices())
        with self.assertRaisesRegex(RuntimeError,"Replacement cancelled"):
            admin.import_invoices([{"invoice_number":"broken","invoice_date":"not a date"}], True)
        self.assertEqual(len(admin.invoices()),before)
        self.assertTrue(admin.journal_voucher(voucher["voucher"]["id"]))
        valid={"invoice_number":"new","invoice_date":"02-03-2024","party_name":"Client","kind":"sale","currency":"USD","subtotal":100,"vat":11,"total":111}
        result=admin.import_invoices([valid],True)
        self.assertEqual(result["imported"],1)
        self.assertEqual(len(admin.invoices()),1)
        self.assertTrue(admin.journal_voucher(voucher["voucher"]["id"]))

    def test_invalid_company_rejected_and_user_can_export_backup(self):
        admin=ApiClient(self.url); admin.login("admin","secret12345")
        company=admin.companies()[0]; admin.select_company_year(company["id"],2024)
        admin.save_user({"username":"staff","password":"staff12345","role":"accountant"})
        staff=ApiClient(self.url); staff.login("staff","staff12345"); staff.select_company_year(company["id"],2024)
        name=Path(staff.create_backup()["path"]).name
        self.assertIn(name,[b["name"] for b in staff.backups()])
        self.assertTrue(staff.download_backup(name)["content"].startswith(b"SQLite format 3"))
        staff.select_company_year("wrong-company",2024)
        with self.assertRaisesRegex(RuntimeError,"Company not found"): staff.invoices()

    def test_background_backup_runs_without_application(self):
        created=backup_all(self.database)
        self.assertTrue(created)
        self.assertEqual(backup_all(self.database),[])

    def test_failed_payment_edit_preserves_existing_record(self):
        db=Database(self.database)
        user=db.user_for_token(db.login("admin","secret12345")["token"])["id"]
        party=db.save_party({"kind":"customer","name":"Edit Safety Client"},user)
        original={"kind":"customer_receipt","party_id":party["id"],"payment_date":"03-03-2024","amount":"50","currency":"USD"}
        payment_id=db.add_payment(original,user)
        with self.assertRaises(ValueError):
            db.update_payment(payment_id,{**original,"amount":"0"},user)
        self.assertTrue(any(row["id"]==payment_id for row in db.list_payments()))

    def test_pdf_pages_make_distinct_invoices_and_keep_scans_visible(self):
        from reportlab.pdfgen import canvas
        output=Path(self.folder.name)/"many.pdf"
        pdf=canvas.Canvas(str(output))
        for number in (101,102):
            pdf.drawString(30,750,f"Invoice No: {number}")
            pdf.drawString(30,730,"Date: 12-03-2024")
            pdf.drawString(30,710,"Subtotal 100")
            pdf.drawString(30,690,"VAT 11")
            pdf.drawString(30,670,"Total 111")
            pdf.showPage()
        pdf.showPage(); pdf.save()
        invoices=read_invoice_pdf_pages(output)
        self.assertEqual(len(invoices),3)
        self.assertEqual([row["page_range"] for row in invoices],["Page 1","Page 2","Page 3"])
        self.assertEqual([row["invoice_number"] for row in invoices[:2]],["101","102"])
        self.assertIn("scanned",invoices[2]["notes"])


if __name__=="__main__": unittest.main()
