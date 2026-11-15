import sys
import unittest
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from commission_engine import analyze_workbooks

def invoice_rows(amount, method="روش نقدی"):
    header=[None]*21; header[0]="اطلاعات تراکنش"; header[1]="1405/07/01"; header[7]="INV-1"; header[11]="مشتری"
    product=[None]*21; product[6]="نام کالا"; product[8]="مامور نمونه"; product[9]=amount; product[12]=1; product[13]=abs(amount); product[14]="کالا"
    return [header,product,["اطلاعات پرداخت+"],["مقدار","روش پرداخت"],[amount,method],[amount,"جمع روش های پرداخت"]]

class EngineTests(unittest.TestCase):
    def test_blank_rate_is_zero(self):
        r=analyze_workbooks(invoice_rows(100000000),[[None,"روش نقدی",None]])
        self.assertEqual(r["fee_total"],0); self.assertEqual(r["commission_total"],1000000)
    def test_fee_is_deducted_first(self):
        r=analyze_workbooks(invoice_rows(100000000,"بانک"),[[None,"بانک",10]])
        self.assertEqual(r["fee_total"],10000000); self.assertEqual(r["commission_total"],900000)
    def test_return_reverses_commission(self):
        r=analyze_workbooks(invoice_rows(-100000000,"بانک"),[[None,"بانک",10]])
        self.assertEqual(r["negative_invoice_count"],1); self.assertEqual(r["commission_total"],-900000)
    def test_payment_mismatch_is_reported(self):
        rows=invoice_rows(100000000); rows[-2][0]=90000000; rows[-1][0]=90000000
        r=analyze_workbooks(rows,[[None,"روش نقدی",None]])
        self.assertEqual(r["issue_invoice_count"],1); self.assertIn("برابر نیست",r["issues"][0]["علت_بررسی"])

if __name__=="__main__":
    unittest.main()
