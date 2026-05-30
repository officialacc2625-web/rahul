import json

def generate_export_function():
    return """function exportFutureStoresCSV() {
    if (productData.length === 0) return;
    const selRBM = $('fsRBM').value;
    const selBDM = $('fsBDM').value;
    const selBranch = $('fsBranch').value;

    const invoiceMeta = {};
    productData.forEach(r => {
        if (r.invoice) {
            const m = window.getBranchMeta(r);
            invoiceMeta[r.invoice] = {
                branch: m ? m.origBranch : (r.branch || ''),
                rbm: r.rbm || (m ? m.rbm : 'Unknown'),
                bdm: r.bdm || (m ? m.bdm : 'Unknown'),
                staff: r.staff || 'Unknown'
            };
        }
    });

    const getMeta = (r) => {
        if (r.invoice && invoiceMeta[r.invoice]) return invoiceMeta[r.invoice];
        const m = window.getBranchMeta(r);
        return {
            branch: m ? m.origBranch : (r.branch || ''),
            rbm: r.rbm || (m ? m.rbm : 'Unknown'),
            bdm: r.bdm || (m ? m.bdm : 'Unknown'),
            staff: r.staff || 'Unknown'
        };
    };

    const passFilters = (m) => {
        if (selRBM && m.rbm !== selRBM) return false;
        if (selBDM && m.bdm !== selBDM) return false;
        if (selBranch && m.branch !== selBranch) return false;
        return true;
    };

    const fP = [], fO = [], fA = [], fS = [];
    productData.forEach(r => { const m = getMeta(r); if (passFilters(m)) { r._m = m; fP.push(r); } });
    osgData.forEach(r => { const m = getMeta(r); if (passFilters(m)) { r._m = m; fO.push(r); } });
    amcData.forEach(r => { const m = getMeta(r); if (passFilters(m)) { r._m = m; fA.push(r); } });
    samsungData.forEach(r => { const m = getMeta(r); if (passFilters(m)) { r._m = m; fS.push(r); } });

    const H = {};
    const getObj = (rbm, bdm, branch, staff, prod) => {
        if (!H[rbm]) H[rbm] = {};
        if (!H[rbm][bdm]) H[rbm][bdm] = {};
        if (!H[rbm][bdm][branch]) H[rbm][bdm][branch] = {};
        if (!H[rbm][bdm][branch][staff]) H[rbm][bdm][branch][staff] = {};
        const pName = prod ? prod.toUpperCase().trim() : 'UNKNOWN';
        if (!H[rbm][bdm][branch][staff][pName]) {
            H[rbm][bdm][branch][staff][pName] = { 
                name: pName, pQty:0, pRev:0, oQty:0, oRev:0, aQty:0, aRev:0, sQty:0, sRev:0, lgPQty:0, lgPRev:0, samPQty:0, samPRev:0 
            };
        }
        return H[rbm][bdm][branch][staff][pName];
    };

    const samAllowed = ['AC', 'MICROWAVE OVEN', 'REFRIGERATOR', 'WASHING MACHINE'];
    fP.forEach(r => {
        const prod = r.product || 'Unknown';
        const obj = getObj(r._m.rbm, r._m.bdm, r._m.branch, r._m.staff, prod);
        obj.pQty += (r.qty || 0); obj.pRev += (r.soldPrice || 0);
        const b = (r.brand || '').toUpperCase();
        if (b.includes('LG')) { obj.lgPQty += (r.qty || 0); obj.lgPRev += (r.soldPrice || 0); }
        if (b.includes('SAMSUNG') && samAllowed.includes(prod.toUpperCase().trim())) { obj.samPQty += (r.qty || 0); obj.samPRev += (r.soldPrice || 0); }
    });

    fO.forEach(r => {
        const prod = r.product || 'Unknown';
        const obj = getObj(r._m.rbm, r._m.bdm, r._m.branch, r._m.staff, prod);
        obj.oQty += (r.qty || 0); obj.oRev += (r.soldPrice || 0);
    });

    fA.forEach(r => {
        const prod = r.product || 'Unknown';
        const obj = getObj(r._m.rbm, r._m.bdm, r._m.branch, r._m.staff, prod);
        obj.aQty += (r.qty || 0); obj.aRev += (r.soldPrice || 0);
    });

    fS.forEach(r => {
        const prod = r.product || 'Unknown';
        const obj = getObj(r._m.rbm, r._m.bdm, r._m.branch, r._m.staff, prod);
        obj.sQty += (r.qty || 0); obj.sRev += (r.soldPrice || 0);
    });

    const flatDataAll = [];
    const flatDataFuture = [];
    Object.keys(H).forEach(rbm => {
        Object.keys(H[rbm]).forEach(bdm => {
            Object.keys(H[rbm][bdm]).forEach(branch => {
                const isFuture = branch.toUpperCase().includes('FUTURE');
                Object.keys(H[rbm][bdm][branch]).forEach(staff => {
                    Object.keys(H[rbm][bdm][branch][staff]).forEach(prod => {
                        const d = H[rbm][bdm][branch][staff][prod];
                        const row = { rbm, bdm, branch, staff, product: prod, ...d };
                        flatDataAll.push(row);
                        if (isFuture) flatDataFuture.push(row);
                    });
                });
            });
        });
    });

    const calcConv = (n, d) => (d > 0) ? (n / d * 100) : 0;

    const buildBrandWorkbook = (brandType) => {
        const wb = XLSX.utils.book_new();

        const applyDesignStyles = (ws, merges, isSheet1 = false, groupCol = -1) => {
            if (!ws['!merges']) ws['!merges'] = [];
            merges.forEach(m => ws['!merges'].push(m));
            const range = XLSX.utils.decode_range(ws['!ref']);
            const colWidths = [];
            for (let C = range.s.c; C <= range.e.c; ++C) colWidths[C] = 13;

            const rowTypes = {};
            let isTotalNext = false;
            let altIdx = 0;
            let groupMap = {};
            let curGrp = '';
            let curGrpIdx = -1;

            const sTitle = { font: { bold: true, color: { rgb: "FFFFFF" }, sz: 12 }, fill: { fgColor: { rgb: "0A2240" } }, alignment: { horizontal: "center", vertical: "center" } };
            const sHeader = { font: { bold: true, color: { rgb: "FFE600" }, sz: 10 }, fill: { fgColor: { rgb: "0A2240" } }, alignment: { horizontal: "center", vertical: "center", wrapText: true } };
            const sTotal = { font: { bold: true, color: { rgb: "FFFFFF" }, sz: 11 }, fill: { fgColor: { rgb: "0A2240" } }, alignment: { horizontal: "center", vertical: "center" } };
            const sDataW = { font: { color: { rgb: "000000" }, sz: 10 }, fill: { fgColor: { rgb: "FFFFFF" } }, alignment: { horizontal: "center", vertical: "center" }, border: { top: { style: 'thin', color: { rgb: 'B0C4DE' } }, bottom: { style: 'thin', color: { rgb: 'B0C4DE' } }, left: { style: 'thin', color: { rgb: 'B0C4DE' } }, right: { style: 'thin', color: { rgb: 'B0C4DE' } } } };
            const sDataL = { font: { color: { rgb: "000000" }, sz: 10 }, fill: { fgColor: { rgb: "E9F0F8" } }, alignment: { horizontal: "center", vertical: "center" }, border: { top: { style: 'thin', color: { rgb: 'B0C4DE' } }, bottom: { style: 'thin', color: { rgb: 'B0C4DE' } }, left: { style: 'thin', color: { rgb: 'B0C4DE' } }, right: { style: 'thin', color: { rgb: 'B0C4DE' } } } };

            for (let R = range.s.r; R <= range.e.r; R++) {
                const cellRef0 = XLSX.utils.encode_cell({r: R, c: 0});
                const cell0 = ws[cellRef0];
                const val0 = cell0 ? (typeof cell0.v === 'string' ? cell0.v.toUpperCase() : String(cell0.v)) : '';
                
                let rType = 'data';
                if (!isSheet1) {
                    if (R === 0) rType = 'title';
                    else if (R === 1) rType = 'header';
                    else if (isTotalNext || val0 === 'TOTAL' || val0.includes('OVERALL')) { rType = 'total'; isTotalNext = false; }
                    else if (val0 === '') rType = 'data'; 
                    else {
                        const gcRef = XLSX.utils.encode_cell({r: R, c: groupCol >= 0 ? groupCol : 0});
                        const gcVal = ws[gcRef] ? ws[gcRef].v : '';
                        if (gcVal !== curGrp) { curGrp = gcVal; curGrpIdx++; }
                        groupMap[R] = curGrpIdx;
                        rType = 'data';
                        altIdx++;
                    }
                }

                rowTypes[R] = rType;

                for (let C = range.s.c; C <= range.e.c; ++C) {
                    const cellRef = XLSX.utils.encode_cell({r: R, c: C});
                    if (!ws[cellRef]) ws[cellRef] = { t: 's', v: '' };

                    const valStr = String(ws[cellRef].v);
                    valStr.split('\\n').forEach(l => { if (l.length + 2 > colWidths[C]) colWidths[C] = l.length + 2; });

                    let cType = rType;
                    if (isSheet1) {
                        if (R === 0 || R === 7 || R === 8) cType = 'title';
                        else if (R === 1 || R === 9) cType = 'header';
                        else if (val0 === 'TOTAL') cType = 'total';
                        else cType = 'data';
                    }

                    let style;
                    if (cType === 'title')  style = JSON.parse(JSON.stringify(sTitle));
                    else if (cType === 'header') style = JSON.parse(JSON.stringify(sHeader));
                    else if (cType === 'total') style = JSON.parse(JSON.stringify(sTotal));
                    else {
                        const idx = isSheet1 ? R : (groupCol >= 0 ? (groupMap[R] || 0) : altIdx);
                        style = JSON.parse(JSON.stringify(idx % 2 === 0 ? sDataL : sDataW));
                    }

                    const isLabelCol = (isSheet1 && C === 0) || (!isSheet1 && C <= 1);
                    if (isLabelCol && cType !== 'title') {
                        style.alignment = { ...style.alignment, horizontal: "left" };
                    }

                    ws[cellRef].s = style;

                    if (!isNaN(ws[cellRef].v) && ws[cellRef].v !== '') {
                        let numVal = Number(ws[cellRef].v);
                        
                        let isConv = false;
                        for (let hR = R; hR >= range.s.r; hR--) {
                            if (rowTypes[hR] === 'header') {
                                const hCellRef = XLSX.utils.encode_cell({r: hR, c: C});
                                if (ws[hCellRef] && typeof ws[hCellRef].v === 'string') {
                                    const v = ws[hCellRef].v.toUpperCase();
                                    if (v.includes('CONV') || v.includes('%')) isConv = true;
                                }
                                break;
                            }
                        }

                        ws[cellRef].v = isConv ? numVal : Math.round(numVal);
                        ws[cellRef].t = 'n';
                        ws[cellRef].z = isConv ? "0.00" : "0";
                    }
                }
            }

            ws['!cols'] = colWidths.map(w => ({ wch: Math.min(Math.max(w, 13), 28) }));
            ws['!rows'] = [];
            for (let R = range.s.r; R <= range.e.r; R++) {
                if (rowTypes[R] === 'title') ws['!rows'][R] = { hpt: 22 };
                else if (rowTypes[R] === 'header') ws['!rows'][R] = { hpt: 20 };
                else ws['!rows'][R] = { hpt: 18 };
            }
        };

        const addSheet = (name, aoa, merges, isSheet1 = false, groupCol = -1) => {
            const ws = XLSX.utils.aoa_to_sheet(aoa);
            applyDesignStyles(ws, merges, isSheet1, groupCol);
            XLSX.utils.book_append_sheet(wb, ws, name.substring(0, 31));
            return ws;
        };

        // Aggregations for Overview and RBM Wise
        const prodMap = {};
        const rbmMap = {};
        let tPQ=0, tQ=0, tPR=0, tR=0;
        
        flatDataAll.forEach(d => {
            if (!prodMap[d.product]) prodMap[d.product] = { pQ:0, pR:0, q:0, r:0 };
            if (d.rbm && d.rbm.toUpperCase() !== 'GENERAL') {
                if (!rbmMap[d.rbm]) rbmMap[d.rbm] = { pQ:0, pR:0, q:0, r:0, prods: {} };
                if (!rbmMap[d.rbm].prods[d.product]) rbmMap[d.rbm].prods[d.product] = { pQ:0, pR:0, q:0, r:0 };
            }
            
            let myPQ = 0, myPR = 0, myQ = 0, myR = 0;
            if (brandType === 'OSG') { myPQ = d.pQty; myPR = d.pRev; myQ = d.oQty; myR = d.oRev; }
            else if (brandType === 'LG_AMC') { myPQ = d.lgPQty; myPR = d.lgPRev; myQ = d.aQty; myR = d.aRev; }
            else if (brandType === 'SAMSUNG') { myPQ = d.samPQty; myPR = d.samPRev; myQ = d.sQty; myR = d.sRev; }

            prodMap[d.product].pQ += myPQ; prodMap[d.product].pR += myPR;
            prodMap[d.product].q += myQ; prodMap[d.product].r += myR;
            
            if (d.rbm && d.rbm.toUpperCase() !== 'GENERAL') {
                rbmMap[d.rbm].pQ += myPQ; rbmMap[d.rbm].pR += myPR;
                rbmMap[d.rbm].q += myQ; rbmMap[d.rbm].r += myR;
                rbmMap[d.rbm].prods[d.product].pQ += myPQ; rbmMap[d.rbm].prods[d.product].pR += myPR;
                rbmMap[d.rbm].prods[d.product].q += myQ; rbmMap[d.rbm].prods[d.product].r += myR;
            }
            
            tPQ += myPQ; tPR += myPR; tQ += myQ; tR += myR;
        });

        // 1. OVERVIEW Sheet
        const aoa1 = [];
        let titleName = '';
        let qtyName = '';
        if (brandType === 'OSG') { titleName = 'OSG-OVERVIEW'; qtyName = 'OSG QTY'; }
        if (brandType === 'LG_AMC') { titleName = 'LG-AMC OVERVIEW'; qtyName = 'LG-AMC QTY'; }
        if (brandType === 'SAMSUNG') { titleName = 'SAMSUNG CARE+ OVERVIEW'; qtyName = 'SAMSUNG QTY'; }

        aoa1.push([titleName, '', '', '', '']);
        aoa1.push(['PRODUCT', 'QTY', 'PRICE', 'QTY-CONV', 'VALUE-CONV']);
        
        Object.keys(prodMap).sort().forEach(k => {
            const p = prodMap[k];
            if (p.q > 0 || p.r > 0) aoa1.push([k, p.q, p.r, calcConv(p.q, p.pQ), calcConv(p.r, p.pR)]);
        });
        aoa1.push(['TOTAL', tQ, tR, calcConv(tQ, tPQ), calcConv(tR, tPR)]);
        
        aoa1.push([]);
        aoa1.push([]);
        const botStartR = aoa1.length;
        
        aoa1.push(['RBM-OVERVIEW', '', '', '', '']);
        aoa1.push(['RBM', 'QTY', 'SALE', 'QTY CONV', 'VALUE CONV']);
        
        Object.keys(rbmMap).sort().forEach(r => {
            const d = rbmMap[r];
            aoa1.push([r, d.q, d.r, calcConv(d.q, d.pQ), calcConv(d.r, d.pR)]);
        });
        
        const merges1 = [
            {s:{r:0,c:0}, e:{r:0,c:4}}, 
            {s:{r:botStartR, c:0}, e:{r:botStartR, c:4}}
        ];
        addSheet('OVERVIEW', aoa1, merges1, true);

        // 2. RBM WISE Sheet
        const aoa2 = [['RBM WISE OVERVIEW'], ['RBM', 'Product', 'Product Qty', qtyName, 'Qty Conv%', 'Val Conv%', 'OVERALL Qty Conv%', 'OVERALL Val Conv%']];
        const merges2 = [{s:{r:0,c:0}, e:{r:0,c:7}}];
        let rIdx = 2;
        Object.keys(rbmMap).sort().forEach(r => {
            let startR = rIdx;
            const prods = Object.keys(rbmMap[r].prods).sort();
            prods.forEach(k => {
                const p = rbmMap[r].prods[k];
                aoa2.push([r, k, p.pQ, p.q, calcConv(p.q, p.pQ), calcConv(p.r, p.pR), '', '']);
                rIdx++;
            });
            if (rIdx > startR + 1) merges2.push({s:{r:startR, c:0}, e:{r:rIdx-1, c:0}});
            
            const oQtyConv = calcConv(rbmMap[r].q, rbmMap[r].pQ);
            const oValConv = calcConv(rbmMap[r].r, rbmMap[r].pR);
            aoa2[startR][6] = oQtyConv; aoa2[startR][7] = oValConv;
            
            if (rIdx > startR + 1) {
                merges2.push({s:{r:startR, c:6}, e:{r:rIdx-1, c:6}});
                merges2.push({s:{r:startR, c:7}, e:{r:rIdx-1, c:7}});
            }
        });
        addSheet('RBM WISE', aoa2, merges2, false, 0);

        // 3. STORE WISE Sheet
        const aoa3 = [['FUTURE STORES — STORE WISE'], ['BDM', 'Branch', 'Product', 'Product Qty', qtyName, 'Qty Conv%', 'Val Conv%', 'OVERALL Qty Conv%', 'OVERALL Val Conv%']];
        const merges3 = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 8 } }];
        
        const bdmBranchMap = {};
        flatDataFuture.forEach(d => {
            if (!bdmBranchMap[d.bdm]) bdmBranchMap[d.bdm] = {};
            if (!bdmBranchMap[d.bdm][d.branch]) bdmBranchMap[d.bdm][d.branch] = {};
            if (!bdmBranchMap[d.bdm][d.branch][d.product]) bdmBranchMap[d.bdm][d.branch][d.product] = { pQ:0, pR:0, q:0, r:0 };
            const p = bdmBranchMap[d.bdm][d.branch][d.product];
            
            let myPQ = 0, myPR = 0, myQ = 0, myR = 0;
            if (brandType === 'OSG') { myPQ = d.pQty; myPR = d.pRev; myQ = d.oQty; myR = d.oRev; }
            else if (brandType === 'LG_AMC') { myPQ = d.lgPQty; myPR = d.lgPRev; myQ = d.aQty; myR = d.aRev; }
            else if (brandType === 'SAMSUNG') { myPQ = d.samPQty; myPR = d.samPRev; myQ = d.sQty; myR = d.sRev; }
            
            p.pQ+=myPQ; p.pR+=myPR; p.q+=myQ; p.r+=myR;
        });

        let bdmStart = 2;
        Object.keys(bdmBranchMap).sort().forEach(bdm => {
            let branchStart = bdmStart;
            Object.keys(bdmBranchMap[bdm]).sort().forEach(branch => {
                let pStart = branchStart;
                let t_pQ=0, t_pR=0, t_q=0, t_r=0;
                const prods = Object.keys(bdmBranchMap[bdm][branch]).sort();
                prods.forEach(k => {
                    const p = bdmBranchMap[bdm][branch][k];
                    t_pQ+=p.pQ; t_pR+=p.pR; t_q+=p.q; t_r+=p.r;
                    aoa3.push([bdm, branch, k, p.pQ, p.q, calcConv(p.q, p.pQ), calcConv(p.r, p.pR), '', '']);
                    branchStart++;
                });
                aoa3[pStart][7] = calcConv(t_q, t_pQ); aoa3[pStart][8] = calcConv(t_r, t_pR);
                if (branchStart > pStart + 1) {
                    merges3.push({ s: { r: pStart, c: 1 }, e: { r: branchStart - 1, c: 1 } });
                    merges3.push({ s: { r: pStart, c: 7 }, e: { r: branchStart - 1, c: 7 } });
                    merges3.push({ s: { r: pStart, c: 8 }, e: { r: branchStart - 1, c: 8 } });
                }
            });
            if (branchStart > bdmStart + 1) merges3.push({ s: { r: bdmStart, c: 0 }, e: { r: branchStart - 1, c: 0 } });
            bdmStart = branchStart;
        });
        addSheet('STORE WISE', aoa3, merges3, false, 0);

        // 4. STAFF WISE Sheet
        const aoa4 = [['BRANCH', 'RBM', 'BDM', 'Staff', 'Product', 'Product Qty', qtyName, 'Qty Conv%', 'Val Conv%']];
        const flatSorted = [...flatDataFuture].sort((a,b) => a.branch.localeCompare(b.branch) || a.bdm.localeCompare(b.bdm) || a.staff.localeCompare(b.staff) || a.product.localeCompare(b.product));
        flatSorted.forEach(d => {
            let myPQ = 0, myPR = 0, myQ = 0, myR = 0;
            if (brandType === 'OSG') { myPQ = d.pQty; myPR = d.pRev; myQ = d.oQty; myR = d.oRev; }
            else if (brandType === 'LG_AMC') { myPQ = d.lgPQty; myPR = d.lgPRev; myQ = d.aQty; myR = d.aRev; }
            else if (brandType === 'SAMSUNG') { myPQ = d.samPQty; myPR = d.samPRev; myQ = d.sQty; myR = d.sRev; }
            
            aoa4.push([d.branch, d.rbm, d.bdm, d.staff, d.product, myPQ, myQ, calcConv(myQ, myPQ), calcConv(myR, myPR)]);
        });
        addSheet('STAFF WISE', aoa4, [], false, 2);

        let filename = 'Future_Stores_Report.xlsx';
        if (brandType === 'OSG') filename = 'Future_Stores_OSG.xlsx';
        if (brandType === 'LG_AMC') filename = 'Future_Stores_LG_AMC.xlsx';
        if (brandType === 'SAMSUNG') filename = 'Future_Stores_Samsung.xlsx';
        
        XLSX.writeFile(wb, filename);
    };

    exportBrandWorkbook('OSG');
    exportBrandWorkbook('LG_AMC');
    exportBrandWorkbook('SAMSUNG');
}
"""

with open('c:\\Users\\rahul_myg\\Downloads\\New folder\\app.js', 'r', encoding='utf-8') as f:
    content = f.read()

import re
# Regex to find function exportFutureStoresCSV() { ... }
start_str = "function exportFutureStoresCSV() {"
end_str = "\n    // ---- PRODUCT DETAILS PAGE ----"

start_idx = content.find(start_str)
end_idx = content.find(end_str)

if start_idx != -1 and end_idx != -1:
    new_content = content[:start_idx] + generate_export_function() + content[end_idx:]
    with open('c:\\Users\\rahul_myg\\Downloads\\New folder\\app.js', 'w', encoding='utf-8') as f:
        f.write(new_content)
    print("Successfully replaced exportFutureStoresCSV")
else:
    print("Could not find start or end index", start_idx, end_idx)
