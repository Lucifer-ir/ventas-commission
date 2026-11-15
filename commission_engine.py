"""Ventas commission engine. Runs in-browser through Pyodide; XLSX bytes remain local."""
import json, math, re, unicodedata
from collections import defaultdict

def _text(v):
    if v is None: return ""
    s = unicodedata.normalize("NFKC", str(v)).strip()
    s = s.replace("ي", "ی").replace("ى", "ی").replace("ك", "ک").replace("\u0640", "")
    return re.sub(r"\\s+", " ", s)

def _key(v):
    return _text(v).replace("\u200c", "").replace(" ", "").lower()

def _num(v):
    if v is None or v == "": return None
    try:
        n = float(str(v).replace(",", "").replace("٬", "").replace("٫", ".").strip())
        return n if math.isfinite(n) else None
    except (TypeError, ValueError):
        return None

def _cell(row, i):
    return row[i] if isinstance(row, (list, tuple)) and len(row) > i else None

def parse_rates(rows):
    rates = {}
    for row in rows or []:
        name = _text(_cell(row, 1))
        if not name or name in ("روش پرداخت", "درصد"): continue
        value = _num(_cell(row, 2))
        rates[_key(name)] = 0.0 if value is None else value
    return rates

def _payment_markup_match(payments, rates, product_total, payment_total, tolerance_pp=0.5):
    """Recognize a payment surcharge close to the fee-sheet rate."""
    if product_total <= 0 or payment_total <= product_total:
        return None
    valid = []
    for payment in payments:
        key = _key(payment["method"])
        rate = rates.get(key)
        if rate is not None and rate > 0 and payment["amount"] > 0:
            valid.append((payment["amount"], rate))
    paid = sum(amount for amount, _ in valid)
    if paid <= 0:
        return None
    expected = sum(amount * rate for amount, rate in valid) / paid
    observed = (payment_total / product_total - 1) * 100.0
    delta = abs(observed - expected)
    if delta <= tolerance_pp:
        return {
            "expected_rate": expected, "observed_rate": observed,
            "delta_pp": delta, "tolerance_pp": tolerance_pp,
            "expected_payment_total": product_total * (1 + expected / 100.0),
            "difference_from_expected_payment": payment_total - product_total * (1 + expected / 100.0)
        }
    return None


def parse_invoices(rows, rates=None):
    rows = rows or []
    starts = [i for i, row in enumerate(rows) if _text(_cell(row, 0)) == "اطلاعات تراکنش"]
    result = []
    for pos, start in enumerate(starts):
        end = starts[pos + 1] if pos + 1 < len(starts) else len(rows)
        block, header = rows[start:end], rows[start]
        invoice = {
            "date": _text(_cell(header, 1)),
            "number": _text(_cell(header, 7)),
            "customer": _text(_cell(header, 11)),
            "items": [], "payments": [], "adjustments": [], "amount_estimate": None,
            "amount_inferred": False, "zero_cash_inferred": False, "markup_match": None,
            "blank_quantity_count": 0, "amount_inferred_reason": "", "problems": []
        }
        payment_at = next((i for i, row in enumerate(block)
                           if _text(_cell(row, 0)) == "اطلاعات پرداخت+"), None)
        product_rows = block if payment_at is None else block[:payment_at]
        for row in product_rows:
            if _text(_cell(row, 6)) != "نام کالا": continue
            amount, agent = _num(_cell(row, 9)), _text(_cell(row, 8))
            invoice["items"].append({
                "agent": agent, "amount": amount, "product": _text(_cell(row, 14)),
                "quantity": _num(_cell(row, 12)), "unit_price": _num(_cell(row, 13))
            })
            if not agent: invoice["problems"].append("نام مامور فروش در یک یا چند ردیف خالی است")
        for row in product_rows:
            adjustment_amount = _num(_cell(row, 0))
            adjustment_label = _text(_cell(row, 1))
            if adjustment_amount is not None and adjustment_amount != 0 and adjustment_label:
                invoice["adjustments"].append({"amount": adjustment_amount, "description": adjustment_label})
        if payment_at is None:
            invoice["problems"].append("بخش اطلاعات پرداخت وجود ندارد")
        else:
            for row in block[payment_at + 2:]:
                method = _text(_cell(row, 1))
                if method == "جمع روش های پرداخت": break
                if not method or method == "روش پرداخت": continue
                amount = _num(_cell(row, 0))
                if amount is None:
                    invoice["problems"].append("مبلغ یکی از روش‌های پرداخت خالی یا نامعتبر است")
                else:
                    invoice["payments"].append({"method": method, "amount": amount})
        # Reconstruct missing product totals from unit price x quantity plus
        # signed invoice adjustments. Blank quantities are treated as one and
        # recorded explicitly in the detail report.
        if invoice["items"] and all(item["amount"] is None for item in invoice["items"]):
            if all(item["unit_price"] is not None for item in invoice["items"]):
                gross_lines = []
                for item in invoice["items"]:
                    if item["quantity"] is None:
                        invoice["blank_quantity_count"] += 1
                    qty = item["quantity"] if item["quantity"] is not None else 1.0
                    gross_lines.append(item["unit_price"] * qty)
                gross_total = sum(gross_lines)
                adjustment_total = sum(x["amount"] for x in invoice["adjustments"])
                estimate = gross_total + adjustment_total
                paid_total = sum(x["amount"] for x in invoice["payments"])
                invoice["amount_estimate"] = {
                    "gross": gross_total, "adjustments": adjustment_total,
                    "estimated_total": estimate, "payment_total": paid_total,
                    "difference": estimate - paid_total
                }
                zero_cash_entry = (
                    paid_total == 0
                    and any(_key(payment["method"]) == _key("نقد") and payment["amount"] == 0
                            for payment in invoice["payments"])
                )
                invoice["markup_match"] = _payment_markup_match(
                    invoice["payments"], rates or {}, estimate, paid_total
                )
                # Some source exports use a zero-value cash line as a
                # placeholder. In this exact case, derive the sale from unit
                # price and signed adjustments, rather than from payment zero.
                if gross_total > 0 and (
                    abs(estimate - paid_total) <= 1 or (zero_cash_entry and estimate > 0)
                    or invoice["markup_match"] is not None
                ):
                    for item, gross in zip(invoice["items"], gross_lines):
                        item["amount"] = gross + adjustment_total * (gross / gross_total)
                    invoice["amount_inferred"] = True
                    invoice["zero_cash_inferred"] = zero_cash_entry and abs(estimate - paid_total) > 1
                    invoice["amount_inferred_reason"] = (
                        "تطبیق با نرخ افزایش روش پرداخت" if invoice["markup_match"]
                        else "ردیف نقد با مبلغ صفر" if invoice["zero_cash_inferred"]
                        else "تطبیق مبلغ پرداخت"
                    )
                else:
                    invoice["problems"].append(
                        "مبلغ نهایی کالا در فایل خالی است؛ جمع قیمت واحدها "
                        f"{gross_total:,.0f}، جمع تعدیلات فاکتور {adjustment_total:,.0f}، "
                        f"مبلغ برآوردی پس از تعدیلات {estimate:,.0f} و جمع پرداخت ثبت‌شده "
                        f"{paid_total:,.0f} است؛ اختلاف {estimate-paid_total:,.0f} ریال است. "
                        "به‌دلیل عدم تطبیق، مبلغ خودکار ثبت نشد."
                    )
            else:
                invoice["problems"].append("مبلغ نهایی کالا خالی است و حداقل یک قیمت واحد نیز ثبت نشده؛ مبلغ قابل استنباط نیست")
        elif any(item["amount"] is None for item in invoice["items"]):
            invoice["problems"].append("مبلغ نهایی بعضی ردیف‌های کالا خالی است؛ فاکتور برای جلوگیری از محاسبه ناقص نیاز به بررسی دارد")
        if not invoice["items"]: invoice["problems"].append("هیچ ردیف کالایی شناسایی نشد")
        if not invoice["payments"]: invoice["problems"].append("هیچ روش پرداخت معتبری ثبت نشده است")
        if not invoice["date"]: invoice["problems"].append("تاریخ فاکتور ثبت نشده است")
        if not invoice["number"]: invoice["problems"].append("شماره فاکتور ثبت نشده است")
        result.append(invoice)
    return result

def _round(v):
    return round(float(v or 0), 2)

def analyze_workbooks(sales_rows, rate_rows, commission_percent=1.0, month_filter="all"):
    pct = _num(commission_percent)
    if pct is None or pct < 0:
        raise ValueError("درصد پورسانت باید عددی بزرگ‌تر یا مساوی صفر باشد")
    rates = parse_rates(rate_rows)
    invoices = parse_invoices(sales_rows, rates)
    months = sorted({x["date"][:7] for x in invoices if len(x["date"]) >= 7})
    selected = [x for x in invoices if month_filter == "all" or x["date"].startswith(month_filter)]
    staff = defaultdict(lambda: {"sales":0.0,"fee":0.0,"base":0.0,"commission":0.0,"lines":0,"invoices":set()})
    issues, details, unknown_all = [], [], set()
    sales_total = fee_total = base_total = 0.0
    negative_count = 0
    total_payments = 0.0
    monthly = defaultdict(lambda: {"sales":0.0,"payments":0.0,"fees":0.0,"commission":0.0,"invoices":0})
    payment_methods = defaultdict(lambda: {"amount":0.0,"fee":0.0,"count":0,"rate":0.0})
    installment_match_count = 0
    inferred_count = 0
    for inv in selected:
        problems = list(dict.fromkeys(inv["problems"]))
        item_total = sum((x["amount"] or 0.0) for x in inv["items"])
        payment_total = sum(x["amount"] for x in inv["payments"])
        is_return = item_total < 0
        if is_return: negative_count += 1
        transaction_type = "برگشت / کاهنده" if is_return else "فروش"
        fee = 0.0
        unknown = []
        month_key = inv["date"][:7] if len(inv["date"]) >= 7 else "نامشخص"
        for pay in inv["payments"]:
            k = _key(pay["method"])
            if k not in rates:
                if abs(pay["amount"]) > 0: unknown.append(pay["method"])
            else:
                method_fee = pay["amount"] * rates[k] / 100.0
                fee += method_fee
                method_row = payment_methods[pay["method"]]
                method_row["amount"] += pay["amount"]
                method_row["fee"] += method_fee
                method_row["count"] += 1
                method_row["rate"] = rates[k]
        if not inv["markup_match"]:
            inv["markup_match"] = _payment_markup_match(inv["payments"], rates, item_total, payment_total)
        if (inv["items"] and abs(item_total - payment_total) > 1
                and not inv["zero_cash_inferred"] and not inv["markup_match"]):
            problems.append(f"جمع مبلغ نهایی کالاها ({item_total:,.0f}) با جمع پرداخت‌ها ({payment_total:,.0f}) برابر نیست")
        if unknown:
            unknown_all.update(unknown)
            problems.append("روش پرداخت در فایل درصدها پیدا نشد: " + "، ".join(sorted(set(unknown))))
        # Absolute values make this check valid for both sales and returns.
        if abs(fee) > abs(item_total) + 1:
            problems.append("قدر مطلق کارمزد محاسبه‌شده از قدر مطلق فروش نهایی فاکتور بیشتر است")
        problems = list(dict.fromkeys(problems))
        if problems:
            issues.append({
                "تاریخ":inv["date"],"شماره_فاکتور":inv["number"],"مشتری":inv["customer"],
                "ماموران_فروش":"، ".join(sorted({item["agent"] for item in inv["items"] if item["agent"]})),
                "کالاها":"، ".join(item["product"] for item in inv["items"] if item["product"]),
                "نوع_تراکنش":transaction_type,"فروش_نهایی_کالاها":_round(item_total),
                "جمع_پرداخت":_round(payment_total),"کارمزد_محاسبه‌شده":_round(fee),
                "مبلغ_برآوردی_از_قیمت_واحد":(_round(inv["amount_estimate"]["estimated_total"]) if inv["amount_estimate"] else None),
                "اختلاف_برآورد_با_پرداخت":(_round(inv["amount_estimate"]["difference"]) if inv["amount_estimate"] else None),
                "علت_بررسی":" | ".join(problems)
            })
            details.append({
                "تاریخ":inv["date"],"شماره_فاکتور":inv["number"],"مشتری":inv["customer"],
                "ماموران_فروش":"، ".join(sorted({item["agent"] for item in inv["items"] if item["agent"]})),
                "کالاها":"، ".join(item["product"] for item in inv["items"] if item["product"]),
                "نوع_تراکنش":transaction_type,"فروش_نهایی_کالاها":_round(item_total),
                "جمع_پرداخت":_round(payment_total),"کارمزد":_round(fee),
                "مبلغ_برآوردی":(_round(inv["amount_estimate"]["estimated_total"]) if inv["amount_estimate"] else None),
                "مبنای_پورسانت":0,"پورسانت":0,"وضعیت":"نیازمند بررسی"
            })
            continue
        invoice_base = item_total - fee
        invoice_commission = invoice_base * pct / 100.0
        sales_total += item_total
        fee_total += fee
        base_total += invoice_base
        total_payments += payment_total
        monthly[month_key]["sales"] += item_total
        monthly[month_key]["payments"] += payment_total
        monthly[month_key]["fees"] += fee
        monthly[month_key]["commission"] += invoice_commission
        monthly[month_key]["invoices"] += 1
        if inv["markup_match"]: installment_match_count += 1
        if inv["amount_inferred"]: inferred_count += 1
        for item in inv["items"]:
            amount, agent = item["amount"] or 0.0, item["agent"]
            if not agent: continue
            share = amount / item_total if item_total else 0.0
            allocated_fee = fee * share
            employee_base = amount - allocated_fee
            row = staff[agent]
            row["sales"] += amount
            row["fee"] += allocated_fee
            row["base"] += employee_base
            row["commission"] += employee_base * pct / 100.0
            row["lines"] += 1
            row["invoices"].add(inv["number"])
        details.append({
            "تاریخ":inv["date"],"شماره_فاکتور":inv["number"],"مشتری":inv["customer"],
            "نوع_تراکنش":transaction_type,"فروش_نهایی_کالاها":_round(item_total),
            "جمع_پرداخت":_round(payment_total),"روش_پرداخت":"، ".join(pay["method"] for pay in inv["payments"]),
            "تعداد_ردیف_کالا":len(inv["items"]),"کارمزد":_round(fee),
            "مبنای_پورسانت":_round(invoice_base),"پورسانت":_round(invoice_commission),
            "مبلغ_نهایی_استنباط_شده":("بله - " + inv.get("amount_inferred_reason", "بازسازی از قیمت واحد")) if inv["amount_inferred"] else "خیر",
            "نرخ_افزایش_روش_پرداخت":_round(inv["markup_match"]["expected_rate"]) if inv["markup_match"] else None,
            "افزایش_مشاهده‌شده_درصد":_round(inv["markup_match"]["observed_rate"]) if inv["markup_match"] else None,
            "اختلاف_با_نرخ_مجاز_واحد_درصد":_round(inv["markup_match"]["delta_pp"]) if inv["markup_match"] else None,
            "تعداد_خالی_فرض‌شده_یک":inv["blank_quantity_count"],
            "هشدار_پرداخت":"روش نقد با مبلغ صفر ثبت شده؛ فروش از قیمت واحد و تخفیف محاسبه شد" if inv["zero_cash_inferred"] else "",
            "وضعیت_تطبیق_اقساط":"تطبیق با نرخ روش پرداخت تایید شد" if inv["markup_match"] else "",
            "وضعیت":"محاسبه شد"
        })
    staff_rows = []
    for name, row in sorted(staff.items(), key=lambda pair: pair[1]["commission"], reverse=True):
        staff_rows.append({
            "کارمند":name,"تعداد_فاکتور":len(row["invoices"]),"تعداد_ردیف_کالا":row["lines"],
            "فروش_نهایی":_round(row["sales"]),"سهم_کارمزد":_round(row["fee"]),
            "فروش_پس_از_کارمزد":_round(row["base"]),"درصد_پورسانت":pct,
            "پورسانت":_round(row["commission"])
        })
    computed = sum(1 for x in details if x["وضعیت"] == "محاسبه شد")
    staff_lookup = {row["کارمند"]: row for row in staff_rows}
    monthly_rows = [
        {"ماه": key, "فروش": _round(value["sales"]), "پرداخت": _round(value["payments"]),
         "کارمزد": _round(value["fees"]), "پورسانت": _round(value["commission"]),
         "تعداد_فاکتور": value["invoices"]}
        for key, value in sorted(monthly.items())
    ]
    payment_rows = [
        {"روش_پرداخت": name, "مبلغ_پرداخت": _round(value["amount"]),
         "کارمزد": _round(value["fee"]), "درصد_کارمزد": value["rate"],
         "تعداد_استفاده": value["count"]}
        for name, value in sorted(payment_methods.items(), key=lambda item: item[1]["amount"], reverse=True)
    ]
    return {
        "months":months,"rates_loaded":len(rates),"invoice_count_total":len(invoices),
        "invoice_count_filtered":len(selected),"computed_invoice_count":computed,
        "issue_invoice_count":len(issues),"negative_invoice_count":negative_count,
        "zero_cash_inferred_count":sum(1 for x in selected if x["zero_cash_inferred"]),
        "inferred_count":inferred_count,
        "installment_markup_reconciled_count":installment_match_count,
        "sales_total":_round(sales_total),"payment_total":_round(total_payments),
        "fee_total":_round(fee_total),"commission_base_total":_round(base_total),
        "commission_total":_round(sum(x["پورسانت"] for x in staff_rows)),
        "average_invoice":_round(sales_total / computed) if computed else 0,
        "reconciliation_rate":_round(computed * 100 / len(selected)) if selected else 0,
        "unknown_methods":sorted(unknown_all),"staff":staff_rows,"issues":issues,"invoices":details,
        "monthly":monthly_rows,"payment_methods":payment_rows,
        "fee_rule":"درصد خالی = صفر درصد کارمزد؛ افزایش اقساطی نزدیک به نرخ فایل با تلورانس ۰٫۵ واحد درصد پذیرفته می‌شود"
    }

def analyze_workbooks_json(sales_json, rates_json, commission_percent=1.0, month_filter="all"):
    result = analyze_workbooks(json.loads(sales_json), json.loads(rates_json), commission_percent, month_filter)
    return json.dumps(result, ensure_ascii=False, separators=(",", ":"))
