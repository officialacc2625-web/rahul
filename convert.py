"""
Convert large XLSX files to optimized CSV for the Analytics Portal.
Only exports columns needed by the portal to keep file size small.

Usage: python convert.py "path/to/file.xlsx"
Output: creates "(filename) (portal).csv" in the same folder.
"""
import sys, os, csv, openpyxl

PORTAL_COLUMNS = [
    'Slno','Branch','RBM','BDM','Staff','Product','Category','Brand',
    'Invoice Number','QTY','MOP','Sold Price','Taxable Value','Tax',
    'Processing Charge','Service Charge','DBD Charge',
    'Indirect Discount','Direct Discount','Addition','Deduction',
    'Store Name','Store Code','Store','Quantity','Invoice No',
    'Plan Price','EWS Qty','Net Qty','Net Prod Qty',
]

if len(sys.argv) < 2:
    print("Usage: python convert.py <file.xlsx>")
    sys.exit(1)

xlsx_path = sys.argv[1]
if not os.path.exists(xlsx_path):
    print(f"File not found: {xlsx_path}")
    sys.exit(1)

base = os.path.splitext(xlsx_path)[0]
csv_path = base + ' (portal).csv'
print(f"Converting: {xlsx_path}")

wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
total = 0

with open(csv_path, 'w', newline='', encoding='utf-8') as f:
    writer = None
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        print(f"  Sheet: {sheet_name}...")
        row_num = 0
        headers = []
        keep_idx = []
        for row in ws.rows:
            vals = [c.value if c.value is not None else '' for c in row]
            if row_num == 0:
                headers = [str(v).strip() for v in vals]
                if writer is None:
                    # Filter to only portal-needed columns (case-insensitive match)
                    portal_lower = {p.lower() for p in PORTAL_COLUMNS}
                    keep_idx = [i for i, h in enumerate(headers) if h.lower() in portal_lower]
                    if not keep_idx:
                        keep_idx = list(range(len(headers)))  # Keep all if no match
                    out_headers = [headers[i] for i in keep_idx]
                    writer = csv.writer(f)
                    writer.writerow(out_headers)
            else:
                writer.writerow([vals[i] if i < len(vals) else '' for i in keep_idx])
                total += 1
            row_num += 1
            if total % 50000 == 0 and total > 0:
                print(f"    ...{total:,} rows")

wb.close()
size = os.path.getsize(csv_path) / (1024*1024)
print(f"\nDone! {total:,} rows -> {csv_path} ({size:.1f} MB)")
print("Upload this CSV file in the portal.")
