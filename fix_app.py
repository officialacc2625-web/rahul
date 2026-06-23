import re

with open('app.js', 'r', encoding='utf-8') as f:
    content = f.read()

# I know I broke the area starting from document.addEventListener('click' to the end of enderLowConvPage block
# Let's find exactly what's there and replace it with the correct implementation.

# In the current file, we have a broken uildStaffStats because my replacement overwrote it:
# The replacement started with:
#     if (inv) {
#         if (selectedProducts && !selectedProducts.includes(inv.product)) return;
#         s = inv.staff;
# And went until the end of enderLowConvPage.

# First, I will replace the broken section starting from document.addEventListener('click', function(e) { 
# up to the start of unction exportLowConvCSV() { 
# using a regex.

# We find where document.addEventListener('click' is, wait, there are multiple. 
# It's specifically the one for closing dropdowns.

# To be safe, let's just write the EXACT correct code for this whole block.
# We will locate the line: // Close dropdown on click outside
# and the line     function exportLowConvCSV() {

start_idx = content.find('// Close dropdown on click outside')
end_idx = content.find('    function exportLowConvCSV() {')

if start_idx != -1 and end_idx != -1:
    correct_code = '''// Close dropdown on click outside
    document.addEventListener('click', function(e) {
        const wrapper = document.getElementById('lcProductMultiWrapper');
        if (wrapper && !wrapper.contains(e.target)) {
            const dropdown = document.getElementById('lcProductDropdown');
            if(dropdown) dropdown.style.display = 'none';
        }

        const lcBranchWrapper = document.getElementById('lcBranchMultiWrapper');
        if (lcBranchWrapper && !lcBranchWrapper.contains(e.target)) {
            const bDropdown = document.getElementById('lcBranchDropdown');
            if (bDropdown) bDropdown.style.display = 'none';
        }

        const tcBranchWrapper = document.getElementById('tcBranchMultiWrapper');
        if (tcBranchWrapper && !tcBranchWrapper.contains(e.target)) {
            const tcDropdown = document.getElementById('tcBranchDropdown');
            if (tcDropdown) tcDropdown.style.display = 'none';
        }
        
        const osgQtyWrapper = document.getElementById('lcOsgQtyMultiWrapper');
        if (osgQtyWrapper && !osgQtyWrapper.contains(e.target)) {
            const dropdown = document.getElementById('lcOsgQtyDropdown');
            if(dropdown) dropdown.style.display = 'none';
        }
    });

    // ---- LOW CONV STAFF LOGIC ----
    function buildStaffStats(selectedProducts = null) {
        const invoiceData = {};
        productData.forEach(r => { if (r.invoice) invoiceData[r.invoice] = { staff: r.staff || 'Unknown', product: r.category }; });

        const pByStaff = {};
        productData.forEach(r => {
            if (selectedProducts && !selectedProducts.includes(r.category)) return;
            const s = r.staff || 'Unknown';
            if (!pByStaff[s]) pByStaff[s] = { branch: r.branch, rbm: r.rbm, bdm: r.bdm, rows: [] };
            pByStaff[s].rows.push(r);
        });

        const oByStaff = {};
        osgData.forEach(r => {
            let s = null, pName = null;
            const inv = r.invoice ? invoiceData[r.invoice] : null;
            if (inv) { s = inv.staff; pName = inv.product; }
            else if (r.staff) { s = r.staff; }
            if (!s) return;
            if (selectedProducts && pName && !selectedProducts.includes(pName)) return;
            if (!oByStaff[s]) oByStaff[s] = [];
            oByStaff[s].push(r);
        });

        const lgByStaff = {};
        amcData.forEach(r => {
            let s = null;
            const inv = r.invoice ? invoiceData[r.invoice] : null;
            if (inv) {
                if (selectedProducts && !selectedProducts.includes(inv.product)) return;
                s = inv.staff;
            } else if (!selectedProducts && r.staff) {
                s = r.staff;
            }
            if (!s) return;
            lgByStaff[s] = (lgByStaff[s] || 0) + (r.qty || 0);
        });

        const samByStaff = {};
        samsungData.forEach(r => {
            let s = null;
            const inv = r.invoice ? invoiceData[r.invoice] : null;
            if (inv) {
                if (selectedProducts && !selectedProducts.includes(inv.product)) return;
                s = inv.staff;
            } else if (!selectedProducts && r.staff) {
                s = r.staff; 
            }
            if (!s) return;
            samByStaff[s] = (samByStaff[s] || 0) + (r.qty || 0);
        });

        const allStaff = new Set([...Object.keys(pByStaff), ...Object.keys(oByStaff)]);
        const finalStats = Array.from(allStaff).map(name => {
            const pInfo = pByStaff[name] || { branch: '', rbm: '', bdm: '', rows: [] };
            const oRows = oByStaff[name] || [];
            const pQty = pInfo.rows.reduce((s, r) => s + r.qty, 0);
            const oQty = oRows.reduce((s, r) => s + r.qty, 0);
            const lgOsgQty = lgByStaff[name] || 0;
            const samsungOsgQty = samByStaff[name] || 0;
            const pRev = pInfo.rows.reduce((s, r) => s + r.soldPrice, 0);
            const oRev = oRows.reduce((s, r) => s + r.soldPrice, 0);
            const qtyConv = pQty > 0 ? (oQty / pQty) * 100 : 0;
            const valConv = pRev > 0 ? (oRev / pRev) * 100 : 0;
            return { name, branch: pInfo.branch, rbm: pInfo.rbm, bdm: pInfo.bdm, pQty, oQty, lgOsgQty, samsungOsgQty, pRev, oRev, qtyConv, valConv };
        });
        window.portalStaffStats = finalStats;
        return finalStats;
    }

    function renderLowConvPage() {
        if (productData.length === 0) {
            lcTableWrapper.innerHTML = noDataHTML('Upload data and generate reports first.');
            lcKpiRow.innerHTML = '';
            lcCount.textContent = '0 staff';
            return;
        }

        const minQty = parseFloat(lcMinQty.value) || 0;
        const maxConv = parseFloat(lcMaxConv.value);
        const selectedOsgQty = window.getLcSelectedOsgQty();
        const selectedBranches = window.getLcSelectedBranches();
        const selRBM = lcRBM.value;
        const selBDM = lcBDM.value;

        const selectedProducts = window.getLcSelectedProducts();
        const allStats = buildStaffStats(selectedProducts);

        const branchSet = [...new Set(allStats.map(s => s.branch).filter(Boolean))].sort();
        const rbmSet = [...new Set(allStats.map(s => s.rbm).filter(Boolean))].sort();
        const bdmSet = [...new Set(allStats.map(s => s.bdm).filter(Boolean))].sort();

        const branchDrop = lcBranchDropdown;
        const rbmEl = lcRBM;
        const bdmEl = lcBDM;
        const prevRBM = selRBM;
        const prevBDM = selBDM;

        if (branchDrop) {
            let bHtml = <label style="display: flex; align-items: center; gap: 8px; padding: 4px 0; cursor: pointer; color: var(--text-primary); text-transform: none; font-weight: 500; font-size: 0.85rem; margin:0;"><input type="checkbox" value="ALL"  onchange="window.toggleAllLcBranches(this)"> <strong>All Branches</strong></label><hr style="margin: 4px 0; border: none; border-top: 1px solid var(--border);">;
            bHtml += branchSet.map(b => <label style="display: flex; align-items: center; gap: 8px; padding: 4px 0; cursor: pointer; color: var(--text-primary); text-transform: none; font-weight: 500; font-size: 0.85rem; margin:0;"><input type="checkbox" value="" class="lc-branch-cb"  onchange="window.updateLcBranchLabel()"> </label>).join('');
            branchDrop.innerHTML = bHtml;
            window.updateLcBranchLabel();
        }
        
        rbmEl.innerHTML = '<option value="">All RBMs</option>' +
            rbmSet.map(r => <option value="" ></option>).join('');
        bdmEl.innerHTML = '<option value="">All BDMs</option>' +
            bdmSet.map(b => <option value="" ></option>).join('');

        const filtered = allStats
            .filter(s => s.pQty >= minQty && s.qtyConv <= maxConv)
            .filter(s => { if (!selectedOsgQty) return true; const q = s.oQty; if (q >= 5 && selectedOsgQty.includes('5+')) return true; return selectedOsgQty.includes(String(q)); })
            .filter(s => !selectedBranches || selectedBranches.includes(s.branch))
            .filter(s => !selRBM || s.rbm === selRBM)
            .filter(s => !selBDM || s.bdm === selBDM)
            .sort((a, b) => {
                if (a.qtyConv !== b.qtyConv) return a.qtyConv - b.qtyConv;
                return b.pQty - a.pQty;
            });

        lcCount.textContent = ${filtered.length} staff;

        const totalPQty = filtered.reduce((s, r) => s + r.pQty, 0);
        const totalOQty = filtered.reduce((s, r) => s + r.oQty, 0);
        const totalPRev = filtered.reduce((s, r) => s + r.pRev, 0);
        const zeroConvCount = filtered.filter(r => r.qtyConv === 0).length;
        lcKpiRow.innerHTML = 
            <div class="lc-kpi"><span class="lc-kpi-label">Zero Conv Staff</span><span class="lc-kpi-val loss-text"></span></div>
            <div class="lc-kpi"><span class="lc-kpi-label">Total Product Qty</span><span class="lc-kpi-val"></span></div>
            <div class="lc-kpi"><span class="lc-kpi-label">Total OSG Qty (Sold)</span><span class="lc-kpi-val"></span></div>
            <div class="lc-kpi"><span class="lc-kpi-label">Opportunity Missed (Qty)</span><span class="lc-kpi-val loss-text"></span></div>
            <div class="lc-kpi"><span class="lc-kpi-label">Total Product Revenue</span><span class="lc-kpi-val"></span></div>
        ;

        if (filtered.length === 0) {
            lcTableWrapper.innerHTML = noDataHTML(No staff found with >= product qty and <=% qty conversion.);
            return;
        }

        let html = <table class="data-table">
            <thead><tr>
                <th>#</th><th>Staff</th><th>Branch</th><th>RBM</th><th>BDM</th>
                <th>Prod Qty</th><th>OSG Qty</th><th style="color:#10b981;">LG</th><th style="color:#f59e0b;">SAMSUNG</th><th>Qty Conv%</th><th>Val Conv%</th><th>Prod Rev</th>
            </tr></thead><tbody>;

        filtered.forEach((e, i) => {
            const convCls = e.qtyConv === 0 ? 'loss-val' : (e.qtyConv < 5 ? 'conv-warn' : 'conv-val');
            const rank = i + 1;
            const rankBadge = rank <= 3 ? <span class="rank-badge rank-"></span> : <span class="rank-num"></span>;
            const dlIcon = <button onclick="window.downloadStaffDetails('', '')" title="Download Staff Details" style="background:none;border:none;cursor:pointer;color:var(--primary);padding:0;margin-left:8px;vertical-align:middle;display:inline-flex;align-items:center;"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg></button>;
            html += <tr>
                <td class="number-cell"></td>
                <td style="white-space:nowrap;"><strong></strong></td>
                <td></td>
                <td></td>
                <td></td>
                <td class="number-cell"><strong></strong></td>
                <td class="number-cell"></td>
                <td class="number-cell" style="color:#10b981;font-weight:600;"></td>
                <td class="number-cell" style="color:#f59e0b;font-weight:600;"></td>
                <td class="number-cell ">%</td>
                <td class="number-cell conv-val">%</td>
                <td class="number-cell"></td>
            </tr>;
        });

        html += '</tbody></table>';
        lcTableWrapper.innerHTML = html;
    }
'''
    new_content = content[:start_idx] + correct_code + '\n' + content[end_idx:]
    with open('app.js', 'w', encoding='utf-8') as f:
        f.write(new_content)
    print("Fixed!")
else:
    print("Could not find start/end indices")
