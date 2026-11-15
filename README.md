# Ventas — Commission Calculator

Persian RTL commission calculator for the store's Excel invoice and payment-fee exports.

## GitHub Pages
Open: https://lucifer-ir.github.io/ventas-commission/

The browser uses SheetJS to decode XLSX cells only. Invoice parsing, payment-method matching, fee calculation, reconciliation, fee allocation, negative/refund handling and employee commission calculations run in Python through Pyodide. The Python engine is in `commission_engine.py`.

## Use
1. Select the invoice workbook.
2. Select the payment-fee percentage workbook.
3. Run the analysis; choose a month and set the commission percentage (default 1%).
4. Export employee totals, invoice issues, or invoice details.

## Calculation rules
- Product amounts are read from the invoice report's final product-total column.
- Payment fees are calculated per payment method and payment amount.
- Blank fee percentage = 0% fee, as instructed.
- Multiple-payment fees are allocated proportionally across product rows.
- Negative invoices/returns reverse employee commission rather than being marked as an excessive fee.
- Missing product amounts, missing salesperson/payment data, unknown payment methods, and payment/product total mismatches are excluded from confirmed commissions and shown in the exception report.

## Privacy
Excel files are processed in the browser and are not uploaded to a server. Pyodide and SheetJS are loaded from versioned CDNs, so an internet connection is required.

## Tests
Run from the repository root:
`python -m unittest -v tests/test_commission_engine.py`


## Missing line-total handling
Some invoice exports contain a blank final product amount while the unit price and invoice adjustment rows are present. Python calculates a diagnostic estimate from unit price × quantity (blank quantity is treated as one product row) plus signed invoice adjustments. It only uses the estimate in commission calculations when it reconciles to the recorded payment total within 1 rial. Otherwise the invoice remains excluded from confirmed commissions and the estimate plus difference are shown in the issue report.
