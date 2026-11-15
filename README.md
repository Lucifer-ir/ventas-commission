# Ventas — Commission Calculator

A browser-based Persian RTL tool for calculating salesperson commissions from the store's Excel exports.

## Use
1. Open the GitHub Pages site: https://Lucifer-ir.github.io/ventas-commission/
2. Select the invoice Excel file.
3. Select the payment-fee percentage workbook.
4. Run the calculation, choose a month, and set the employee commission rate (defaults to 1%).
5. Export employee totals, invoice details, or the separate exception report.

## Calculation rules
- Commission base is the invoice's final product amount after payment-method fees.
- Payment-method fees are calculated from each payment amount using the percentage workbook.
- Blank percentages in the rate workbook mean **0% fee**, per the operator's instruction.
- Invoices with missing amounts/salesperson fields, unknown payment method names, or payment/product total mismatches are placed in the exception report and excluded from confirmed commission totals.
- For invoices paid using multiple payment methods, fees are allocated proportionally across product rows and their salespeople.

## Privacy
Excel files are processed in the browser and are not uploaded to GitHub or any server. The XLSX library is loaded from jsDelivr, so an internet connection is needed to use the tool.
