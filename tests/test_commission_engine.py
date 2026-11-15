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


    def test_blank_total_is_reconstructed_only_when_payment_reconciles(self):
        header=[None]*21
        header[0]="اطلاعات تراکنش"; header[1]="1405/07/01"; header[7]="INV-INFER"; header[11]="مشتری نمونه"
        product=[None]*21
        product[6]="نام کالا"; product[8]="مامور نمونه"; product[9]=None
        product[12]=None; product[13]=41860000; product[14]="کالا"
        discount=[-8372000,"جشنواره خرید نقدی"]
        rows=[header,product,discount,["اطلاعات پرداخت+"],["مقدار","روش پرداخت"],[33488000,"پرداخت نقدی"],[33488000,"جمع روش های پرداخت"]]
        result=analyze_workbooks(rows,[[None,"پرداخت نقدی",None]])
        self.assertEqual(result["computed_invoice_count"],1)
        self.assertEqual(result["sales_total"],33488000)
        self.assertEqual(result["commission_total"],334880)
        self.assertEqual(result["invoices"][0]["مبلغ_نهایی_استنباط_شده"],"بله - با تطبیق پرداخت")

    def test_zero_cash_payment_uses_reconstructed_sale_and_keeps_audit_warning(self):
        header=[None]*21
        header[0]="اطلاعات تراکنش"; header[1]="1405/07/01"; header[7]="3-2-20260923-100012"; header[11]="طاهره بیاتی"
        product=[None]*21
        product[6]="نام کالا"; product[8]="69 سعید اورنگ راد"; product[9]=None
        product[12]=None; product[13]=41860000; product[14]="شلوار جین"
        discount=[-8372000,"جشنواره خرید نقدی"]
        rows=[header,product,discount,["اطلاعات پرداخت+"],["مقدار","روش پرداخت"],[0,"نقد"],[0,"جمع روش های پرداخت"]]
        result=analyze_workbooks(rows,[[None,"نقد",None]])
        self.assertEqual(result["computed_invoice_count"],1)
        self.assertEqual(result["issue_invoice_count"],0)
        self.assertEqual(result["sales_total"],33488000)
        self.assertEqual(result["fee_total"],0)
        self.assertEqual(result["commission_total"],334880)
        self.assertEqual(result["zero_cash_inferred_count"],1)
        self.assertEqual(result["staff"][0]["کارمند"],"69 سعید اورنگ راد")
        self.assertIn("روش نقد با مبلغ صفر",result["invoices"][0]["هشدار_پرداخت"])

if __name__=="__main__":
    unittest.main()
