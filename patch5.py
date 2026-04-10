import re

with open('app.js', 'r', encoding='utf-8') as f:
    text = f.read()

# Update export to include Called By and Follow-up
old_export = """        const hdr = ['#', 'Invoice No', 'Customer Name', 'Customer No', 'Staff', 'Branch', 'RBM', 'BDM', 'Product', 'Qty', 'Sold Price'];
        const data = missedUnique.map((r, i) => [
            i + 1, r.invoice, r.customerName || '', r.customerNo || '',
            r.staff || '', r.branch || '', r.rbm || '', r.bdm || '',
            r.product || '', r.qty, Math.round(r.soldPrice)
        ]);"""

new_export = """        const hdr = ['#', 'Invoice No', 'Date', 'Customer Name', 'Customer No', 'Call Status', 'Interest', 'Follow-up Date', 'Called By', 'Remarks', 'Branch', 'Product', 'Sold Price'];
        const data = missedUnique.map((r, i) => {
            const st = coStatusMap[r.invoice || ''] || {};
            let dStr = '';
            if (r.invoiceDate) dStr = new Date(r.invoiceDate).toLocaleDateString();
            else if (r.time) dStr = new Date(r.time).toLocaleDateString();
            
            return [
                i + 1, r.invoice, dStr, r.customerName || '', r.customerNo || '',
                st.callStatus || '', st.interest || '', st.followUpDate || '', st.calledBy || '', st.remarks || '',
                r.branch || '', r.product || '', Math.round(r.soldPrice||0)
            ];
        });"""

text = text.replace(old_export, new_export)

with open('app.js', 'w', encoding='utf-8') as f:
    f.write(text)
