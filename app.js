// ============================================================
// Analytics Portal — Conversion Reports
// Dual-file: Product Data + OSG Data
// Value Conversion = OSG Sold Price / Product Sold Price
// Qty Conversion   = OSG Quantity  / Product Quantity
// ============================================================

(function () {
    'use strict';

    // ---- STATE ----
    let productData = [];      // Parsed rows from Product file
    let osgData = [];          // Parsed rows from OSG file
    let amcData = [];          // Parsed rows from LG AMC file (optional)
    let allData = [];          // Product + AMC merged (for profit/loss)
    let filteredProduct = [];  // After filters
    let filteredOSG = [];      // After filters
    let filteredAll = [];      // After filters (product+amc)
    let chartInstances = {};

    // ---- COLUMN MAPPING (Product / AMC file) ----
    const PRODUCT_COL_MAP = {
        branch: ['branch', 'store name', 'store', 'branch name'],
        rbm: ['rbm', 'rbm name', 'region', 'regional manager'],
        bdm: ['bdm', 'bdm name', 'business development manager'],
        staff: ['staff', 'staff name', 'salesperson', 'sales person', 'employee'],
        product: ['product', 'product name', 'product type'],
        category: ['category', 'item category', 'item group'],
        brand: ['brand', 'brand name'],
        soldPrice: ['sold price', 'soldprice', 'selling price', 'sale price', 'mop'],
        taxableVal: ['taxable value', 'taxable', 'taxable amount'],
        tax: ['tax', 'tax amount', 'gst', 'tax value'],
        qty: ['qty', 'quantity', 'qnty', 'units'],
        discount: ['direct discount', 'discount', 'total discount'],
        indDiscount: ['indirect discount'],
        dbdCharge: ['dbd charge', 'dbd'],
        procCharge: ['processing charge'],
        svcCharge: ['service charge'],
        addition: ['addition'],
        deduction: ['deduction'],
        invoice: ['invoice number', 'invoice no', 'invoice', 'bill no'],
    };

    // ---- COLUMN MAPPING (OSG file) ----
    const OSG_COL_MAP = {
        branch: ['store name', 'store', 'branch', 'branch name'],
        storeCode: ['store code'],
        product: ['product', 'product name', 'product type'],
        category: ['category'],
        brand: ['brand'],
        soldPrice: ['sold price', 'soldprice', 'plan price', 'selling price'],
        qty: ['quantity', 'qty', 'ews qty', 'qnty'],
        invoice: ['invoice no', 'invoice number', 'invoice', 'bill no'],
    };

    // ---- DOM REFERENCES ----
    const $ = id => document.getElementById(id);
    const sidebar = $('sidebar');
    const menuToggle = $('menuToggle');
    const loadingOverlay = $('loadingOverlay');
    const fileCountBadge = $('fileCountBadge');
    const fileCountText = $('fileCountText');
    const btnReset = $('btnReset');
    const btnGenerate = $('btnGenerate');

    // Upload zones
    const uploadZoneProduct = $('uploadZoneProduct');
    const uploadZoneOSG = $('uploadZoneOSG');
    const uploadZoneAMC = $('uploadZoneAMC');
    const fileInputProduct = $('fileInputProduct');
    const fileInputOSG = $('fileInputOSG');
    const fileInputAMC = $('fileInputAMC');
    const productStatus = $('productStatus');
    const osgStatus = $('osgStatus');
    const amcStatus = $('amcStatus');

    // Filters
    const filterRBM = $('filterRBM');
    const filterBranch = $('filterBranch');
    const filterProduct = $('filterProduct');
    const filterBrand = $('filterBrand');
    const filterBDM = $('filterBDM');
    const filterStaff = $('filterStaff');

    // ---- NAVIGATION ----
    document.querySelectorAll('.nav-item').forEach(item => {
        item.addEventListener('click', (e) => {
            e.preventDefault();
            const section = item.dataset.section;
            document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
            item.classList.add('active');
            document.querySelectorAll('.content-section').forEach(s => s.classList.remove('active'));
            $(section).classList.add('active');
            $('pageTitle').textContent = item.querySelector('span').textContent;
            sidebar.classList.remove('open');
        });
    });
    menuToggle.addEventListener('click', () => sidebar.classList.toggle('open'));

    // Report tabs
    document.querySelectorAll('.report-tab').forEach(tab => {
        tab.addEventListener('click', () => {
            document.querySelectorAll('.report-tab').forEach(t => t.classList.remove('active'));
            document.querySelectorAll('.report-panel').forEach(p => p.classList.remove('active'));
            tab.classList.add('active');
            $(tab.dataset.report).classList.add('active');
        });
    });

    // ---- UPLOAD HANDLING ----
    // Product
    setupUploadZone(uploadZoneProduct, fileInputProduct, async (file) => {
        const rows = await parseProductFile(file);
        productData = rows;
        showFileStatus(productStatus, file.name, rows.length);
        checkGenerateReady();
    });

    // OSG
    setupUploadZone(uploadZoneOSG, fileInputOSG, async (file) => {
        const rows = await parseOSGFile(file);
        osgData = rows;
        showFileStatus(osgStatus, file.name, rows.length);
        checkGenerateReady();
    });

    // AMC (optional)
    setupUploadZone(uploadZoneAMC, fileInputAMC, async (file) => {
        const rows = await parseProductFile(file);
        amcData = rows;
        showFileStatus(amcStatus, file.name, rows.length);
    });

    function setupUploadZone(zone, input, onFile) {
        zone.addEventListener('click', () => input.click());
        zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag-over'); });
        zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
        zone.addEventListener('drop', async e => {
            e.preventDefault();
            zone.classList.remove('drag-over');
            if (e.dataTransfer.files.length > 0) {
                showLoading(true);
                try { await onFile(e.dataTransfer.files[0]); } catch (err) { console.error(err); }
                showLoading(false);
            }
        });
        input.addEventListener('change', async () => {
            if (input.files.length > 0) {
                showLoading(true);
                try { await onFile(input.files[0]); } catch (err) { console.error(err); }
                showLoading(false);
            }
        });
    }

    function showFileStatus(el, name, count) {
        el.className = 'upload-status has-data';
        el.innerHTML = `
            <span class="status-icon">✅</span>
            <span class="status-text">${name}</span>
            <span class="status-count">${count} rows</span>
        `;
    }

    function checkGenerateReady() {
        btnGenerate.disabled = !(productData.length > 0 && osgData.length > 0);
    }

    // ---- GENERATE REPORTS BUTTON ----
    btnGenerate.addEventListener('click', () => {
        showLoading(true);
        allData = [...productData, ...amcData];

        fileCountBadge.style.display = 'flex';
        fileCountText.textContent = `${allData.length} product · ${osgData.length} OSG`;
        btnReset.style.display = 'flex';

        populateFilters();
        applyFilters();

        setTimeout(() => {
            document.querySelector('[data-section="dashboard-section"]').click();
            showLoading(false);
        }, 300);
    });

    // ---- PARSING ----
    function parseProductFile(file) {
        return parseExcel(file, PRODUCT_COL_MAP, (row, mapping) => {
            const r = {};
            r.branch = strVal(row, mapping.branch);
            r.rbm = strVal(row, mapping.rbm);
            r.bdm = strVal(row, mapping.bdm);
            r.staff = strVal(row, mapping.staff);
            r.product = strVal(row, mapping.product);
            r.category = strVal(row, mapping.category);
            r.brand = strVal(row, mapping.brand);
            r.invoice = strVal(row, mapping.invoice);
            r.soldPrice = num(getVal(row, mapping.soldPrice, 0));
            r.taxableVal = num(getVal(row, mapping.taxableVal, 0));
            r.tax = num(getVal(row, mapping.tax, 0));
            r.qty = num(getVal(row, mapping.qty, 1)) || 1;
            r.discount = num(getVal(row, mapping.discount, 0));
            r.indDiscount = num(getVal(row, mapping.indDiscount, 0));
            r.dbdCharge = num(getVal(row, mapping.dbdCharge, 0));
            r.procCharge = num(getVal(row, mapping.procCharge, 0));
            r.svcCharge = num(getVal(row, mapping.svcCharge, 0));
            r.addition = num(getVal(row, mapping.addition, 0));
            r.deduction = num(getVal(row, mapping.deduction, 0));

            const totalCost = r.taxableVal + r.discount + r.indDiscount + r.dbdCharge + r.procCharge + r.svcCharge + r.deduction;
            r.revenue = r.soldPrice + r.addition;
            r.profit = r.revenue - totalCost;
            r.isProfit = r.profit >= 0;
            return r;
        });
    }

    function parseOSGFile(file) {
        return parseExcel(file, OSG_COL_MAP, (row, mapping) => {
            const r = {};
            r.branch = strVal(row, mapping.branch);
            r.storeCode = strVal(row, mapping.storeCode);
            r.product = strVal(row, mapping.product);
            r.category = strVal(row, mapping.category);
            r.brand = strVal(row, mapping.brand);
            r.soldPrice = num(getVal(row, mapping.soldPrice, 0));
            r.qty = num(getVal(row, mapping.qty, 1)) || 1;
            r.invoice = strVal(row, mapping.invoice);
            return r;
        });
    }

    function parseExcel(file, colMap, rowMapper) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const wb = XLSX.read(data, { type: 'array', cellDates: true });
                    const sheet = wb.Sheets[wb.SheetNames[0]];
                    const json = XLSX.utils.sheet_to_json(sheet, { defval: '' });
                    if (json.length === 0) { reject(new Error('No data')); return; }
                    const headers = Object.keys(json[0]);
                    const mapping = autoMapColumns(headers, colMap);
                    resolve(json.map(row => rowMapper(row, mapping)));
                } catch (err) { reject(err); }
            };
            reader.onerror = () => reject(new Error('Read error'));
            reader.readAsArrayBuffer(file);
        });
    }

    function autoMapColumns(headers, colMap) {
        const mapping = {};
        for (const [key, aliases] of Object.entries(colMap)) {
            mapping[key] = null;
            // Iterate aliases in order — first alias gets priority
            for (const alias of aliases) {
                const found = headers.find(h => h.toLowerCase().trim() === alias);
                if (found) { mapping[key] = found; break; }
            }
        }
        console.log('[Column Mapping]', JSON.stringify(mapping));
        return mapping;
    }

    function getVal(row, col, def) {
        if (!col) return def;
        const v = row[col];
        return v !== undefined && v !== null && v !== '' ? v : def;
    }
    function strVal(row, col) { return String(getVal(row, col, '') || '').trim(); }
    function num(v) {
        if (typeof v === 'number') return v;
        const n = parseFloat(String(v).replace(/[^0-9.\-]/g, ''));
        return isNaN(n) ? 0 : n;
    }

    // ---- FILTERS ----
    function populateFilters() {
        populateSelect(filterRBM, uniqueVals(allData, 'rbm'), 'All RBMs');
        populateSelect(filterBranch, uniqueVals([...allData, ...osgData], 'branch'), 'All Branches');
        // Product dropdown is predefined — keep as-is, just reset selection
        filterProduct.value = '';
        // Brand populated dynamically from uploaded data
        populateSelect(filterBrand, uniqueVals([...allData, ...osgData], 'brand'), 'All Brands');
        populateSelect(filterBDM, uniqueVals(allData, 'bdm'), 'All BDMs');
        populateSelect(filterStaff, uniqueVals(allData, 'staff'), 'All Staff');
    }

    function uniqueVals(arr, key) {
        const set = new Set();
        arr.forEach(r => { if (r[key]) set.add(r[key]); });
        return Array.from(set).sort();
    }

    function populateSelect(sel, items, defaultLabel) {
        sel.innerHTML = `<option value="">${defaultLabel}</option>`;
        items.forEach(v => { const o = document.createElement('option'); o.value = v; o.textContent = v; sel.appendChild(o); });
    }

    $('btnApplyFilter').addEventListener('click', applyFilters);
    $('btnClearFilter').addEventListener('click', () => {
        filterRBM.value = ''; filterBranch.value = ''; filterProduct.value = '';
        filterBrand.value = ''; filterBDM.value = ''; filterStaff.value = '';
        applyFilters();
    });

    // Auto-apply filters on dropdown change
    [filterRBM, filterBranch, filterProduct, filterBrand, filterBDM, filterStaff].forEach(sel => {
        sel.addEventListener('change', applyFilters);
    });

    function applyFilters() {
        const fRBM = filterRBM.value;
        const fBranch = filterBranch.value;
        const fProduct = filterProduct.value;
        const fBrand = filterBrand.value;
        const fBDM = filterBDM.value;
        const fStaff = filterStaff.value;

        // Filter product/amc data
        filteredAll = allData.filter(r => {
            if (fRBM && r.rbm !== fRBM) return false;
            if (fBranch && r.branch !== fBranch) return false;
            if (fProduct && r.product !== fProduct) return false;
            if (fBrand && r.brand !== fBrand) return false;
            if (fBDM && r.bdm !== fBDM) return false;
            if (fStaff && r.staff !== fStaff) return false;
            return true;
        });

        filteredProduct = productData.filter(r => {
            if (fRBM && r.rbm !== fRBM) return false;
            if (fBranch && r.branch !== fBranch) return false;
            if (fProduct && r.product !== fProduct) return false;
            if (fBrand && r.brand !== fBrand) return false;
            if (fBDM && r.bdm !== fBDM) return false;
            if (fStaff && r.staff !== fStaff) return false;
            return true;
        });

        // Filter OSG data by product, brand, branch (OSG has these fields)
        const hasPersonFilter = fRBM || fBDM || fStaff;

        if (hasPersonFilter) {
            const productInvoices = new Set();
            filteredProduct.forEach(r => { if (r.invoice) productInvoices.add(r.invoice); });

            filteredOSG = osgData.filter(r => {
                if (fBranch && r.branch !== fBranch) return false;
                if (fProduct && r.product !== fProduct) return false;
                if (fBrand && r.brand !== fBrand) return false;
                if (r.invoice && productInvoices.has(r.invoice)) return true;
                if (!r.invoice) return true;
                return false;
            });
        } else {
            filteredOSG = osgData.filter(r => {
                if (fBranch && r.branch !== fBranch) return false;
                if (fProduct && r.product !== fProduct) return false;
                if (fBrand && r.brand !== fBrand) return false;
                return true;
            });
        }

        renderDashboard();
        renderReports();
        renderCharts();
    }

    // ---- CONVERSION CALCULATION ----
    function calcConversion(productRows, osgRows) {
        const pSoldPrice = productRows.reduce((s, r) => s + r.soldPrice, 0);
        const oSoldPrice = osgRows.reduce((s, r) => s + r.soldPrice, 0);
        const pQty = productRows.reduce((s, r) => s + r.qty, 0);
        const oQty = osgRows.reduce((s, r) => s + r.qty, 0);

        return {
            valueConv: pSoldPrice > 0 ? (oSoldPrice / pSoldPrice) * 100 : 0,
            qtyConv: pQty > 0 ? (oQty / pQty) * 100 : 0,
            pSoldPrice, oSoldPrice, pQty, oQty
        };
    }

    // ---- DASHBOARD ----
    function renderDashboard() {
        const data = filteredAll;
        const totalQty = data.reduce((s, r) => s + r.qty, 0);

        const conv = calcConversion(filteredProduct, filteredOSG);

        $('kpiValConv').textContent = conv.valueConv.toFixed(2) + '%';
        $('kpiQtyConv').textContent = conv.qtyConv.toFixed(2) + '%';
        $('kpiQuantity').textContent = formatNumber(totalQty);

        // Conversion breakdown tables
        renderConvTable('convRBMTable', 'rbm');
        renderConvTable('convBDMTable', 'bdm');
        renderConvTable('convStaffTable', 'staff');
        renderConvTable('convBranchTable', 'branch');
        renderConvTable('convProductTable', 'product');
    }

    function renderConvTable(containerId, key) {
        // Group product data
        const pGrouped = groupBy(filteredProduct, key);

        // OSG natively has: branch, product, brand, storeCode, invoice.
        // It does NOT have: rbm, bdm, staff.
        // For native fields group directly; for person-level fields use invoice lookup.
        const OSG_NATIVE_FIELDS = new Set(['branch', 'product', 'brand', 'storeCode', 'invoice']);
        let oGrouped;

        if (OSG_NATIVE_FIELDS.has(key)) {
            oGrouped = groupBy(filteredOSG, key);
        } else {
            // Build invoice → key value lookup from filtered product data
            const invoiceToKey = {};
            filteredProduct.forEach(r => {
                if (r.invoice && r[key]) invoiceToKey[r.invoice] = r[key];
            });

            // Group OSG rows using the invoice lookup
            oGrouped = {};
            filteredOSG.forEach(r => {
                const groupName = (r.invoice && invoiceToKey[r.invoice]) ? invoiceToKey[r.invoice] : null;
                if (groupName) {
                    if (!oGrouped[groupName]) oGrouped[groupName] = [];
                    oGrouped[groupName].push(r);
                }
                // Skip OSG rows that can't be attributed — no "Unknown"
            });
        }

        // Get all unique keys from both (excluding "Unknown")
        const allKeys = new Set([...Object.keys(pGrouped), ...Object.keys(oGrouped)]);
        allKeys.delete('Unknown');
        if (allKeys.size === 0) {
            $(containerId).innerHTML = '<p class="no-data-msg">No data</p>';
            return;
        }

        const entries = Array.from(allKeys).map(name => {
            const pRows = pGrouped[name] || [];
            const oRows = oGrouped[name] || [];
            const pSold = pRows.reduce((s, r) => s + r.soldPrice, 0);
            const oSold = oRows.reduce((s, r) => s + r.soldPrice, 0);
            const pQty = pRows.reduce((s, r) => s + r.qty, 0);
            const oQty = oRows.reduce((s, r) => s + r.qty, 0);
            const valConv = pSold > 0 ? (oSold / pSold) * 100 : 0;
            const qtyConv = pQty > 0 ? (oQty / pQty) * 100 : 0;
            return { name, pSold, oSold, pQty, oQty, valConv, qtyConv };
        }).sort((a, b) => b.pSold - a.pSold);

        let html = `<table class="data-table"><thead><tr>
            <th>${capitalize(key)}</th><th>Prod Rev</th><th>OSG Rev</th><th>Val Conv%</th>
            <th>Prod Qty</th><th>OSG Qty</th><th>Qty Conv%</th>
        </tr></thead><tbody>`;
        entries.forEach(e => {
            html += `<tr>
                <td>${e.name}</td>
                <td class="number-cell">${fmtShort(e.pSold)}</td>
                <td class="number-cell">${fmtShort(e.oSold)}</td>
                <td class="number-cell conv-val">${e.valConv.toFixed(2)}%</td>
                <td class="number-cell">${e.pQty}</td>
                <td class="number-cell">${e.oQty}</td>
                <td class="number-cell conv-val">${e.qtyConv.toFixed(2)}%</td>
            </tr>`;
        });
        html += '</tbody></table>';
        $(containerId).innerHTML = html;
    }

    // ---- REPORTS ----
    function renderReports() {
        renderConversionReport();
        renderFullDataReport();
    }



    function renderConversionReport() {
        // Group by branch — show value and qty conversion
        const pGrouped = groupBy(filteredProduct, 'branch');
        const oGrouped = groupBy(filteredOSG, 'branch');
        const allKeys = new Set([...Object.keys(pGrouped), ...Object.keys(oGrouped)]);

        if (allKeys.size === 0) {
            $('conversionTableWrapper').innerHTML = noDataHTML('Upload both Product and OSG files');
            return;
        }

        const entries = Array.from(allKeys).map(branch => {
            const pRows = pGrouped[branch] || [];
            const oRows = oGrouped[branch] || [];
            const pSold = pRows.reduce((s, r) => s + r.soldPrice, 0);
            const oSold = oRows.reduce((s, r) => s + r.soldPrice, 0);
            const pQty = pRows.reduce((s, r) => s + r.qty, 0);
            const oQty = oRows.reduce((s, r) => s + r.qty, 0);
            const rbms = [...new Set(pRows.map(r => r.rbm).filter(Boolean))].join(', ');
            const bdms = [...new Set(pRows.map(r => r.bdm).filter(Boolean))].join(', ');
            const staffCount = new Set(pRows.map(r => r.staff).filter(Boolean)).size;
            const products = [...new Set([...pRows, ...oRows].map(r => r.product).filter(Boolean))];
            return {
                branch, rbms, bdms, staffCount, productCount: products.length,
                pSold, oSold, pQty, oQty,
                valConv: pSold > 0 ? (oSold / pSold) * 100 : 0,
                qtyConv: pQty > 0 ? (oQty / pQty) * 100 : 0
            };
        }).sort((a, b) => b.pSold - a.pSold);

        let html = `<table class="data-table"><thead><tr>
            <th>Branch</th><th>RBM</th><th>BDM</th><th>Staff</th><th>Products</th>
            <th>Prod Revenue</th><th>OSG Revenue</th><th>Value Conv%</th>
            <th>Prod Qty</th><th>OSG Qty</th><th>Qty Conv%</th>
        </tr></thead><tbody>`;
        entries.forEach(e => {
            html += `<tr>
                <td>${e.branch}</td><td>${e.rbms}</td><td>${e.bdms}</td>
                <td class="number-cell">${e.staffCount}</td><td class="number-cell">${e.productCount}</td>
                <td class="number-cell">${fmtShort(e.pSold)}</td>
                <td class="number-cell">${fmtShort(e.oSold)}</td>
                <td class="number-cell conv-val">${e.valConv.toFixed(2)}%</td>
                <td class="number-cell">${e.pQty}</td>
                <td class="number-cell">${e.oQty}</td>
                <td class="number-cell conv-val">${e.qtyConv.toFixed(2)}%</td>
            </tr>`;
        });
        html += '</tbody></table>';
        $('conversionTableWrapper').innerHTML = html;
    }

    function renderFullDataReport() {
        const rows = filteredAll.slice(0, 500);
        $('fullDataTableWrapper').innerHTML = rows.length === 0 ? noDataHTML('Upload data to view') : buildDetailTable(rows);
    }

    function buildDetailTable(rows) {
        let html = `<table class="data-table"><thead><tr>
            <th>Branch</th><th>RBM</th><th>BDM</th><th>Staff</th><th>Product</th>
            <th>Brand</th><th>Sold Price</th><th>Taxable</th><th>Tax</th><th>Qty</th><th>Profit/Loss</th>
        </tr></thead><tbody>`;
        rows.forEach(r => {
            const cls = r.isProfit ? 'profit-val' : 'loss-val';
            html += `<tr>
                <td>${r.branch}</td><td>${r.rbm}</td><td>${r.bdm}</td><td>${r.staff}</td>
                <td>${r.product}</td><td>${r.brand}</td>
                <td class="number-cell">${fmtShort(r.soldPrice)}</td>
                <td class="number-cell">${fmtShort(r.taxableVal)}</td>
                <td class="number-cell">${fmtShort(r.tax)}</td>
                <td class="number-cell">${r.qty}</td>
                <td class="number-cell ${cls}">${fmtShort(r.profit)}</td>
            </tr>`;
        });
        html += '</tbody></table>';
        return html;
    }

    function noDataHTML(msg) {
        return `<div class="no-data-msg"><svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg><p>${msg}</p></div>`;
    }

    // ---- CHARTS ----
    function renderCharts() {
        renderValConvByBranch();
        renderQtyConvByRBM();
        renderQtyConvByProduct();
        renderProductRevenueChart();
        renderProductPieChart();
    }

    const COLORS = {
        green: 'rgba(16,185,129,0.8)', red: 'rgba(239,68,68,0.8)',
        blue: 'rgba(59,130,246,0.8)', amber: 'rgba(245,158,11,0.8)',
        cyan: 'rgba(6,182,212,0.8)', purple: 'rgba(139,92,246,0.8)',
        orange: 'rgba(249,115,22,0.8)',
    };
    const PALETTE = ['#3b82f6', '#8b5cf6', '#06b6d4', '#10b981', '#f59e0b', '#ef4444', '#ec4899', '#6366f1',
        '#14b8a6', '#f97316', '#a855f7', '#22d3ee', '#84cc16', '#e11d48', '#0ea5e9', '#d946ef'];

    function chartBase(withYCurrency) {
        return {
            responsive: true, maintainAspectRatio: false,
            plugins: { legend: { labels: { color: '#8896b8', font: { family: 'Inter', size: 11 } } } },
            scales: {
                x: { ticks: { color: '#5a6a8a', font: { size: 10 } }, grid: { color: 'rgba(30,42,69,0.5)' } },
                y: { ticks: { color: '#5a6a8a', font: { size: 10 }, callback: withYCurrency ? v => fmtShort(v) : v => v + '%' }, grid: { color: 'rgba(30,42,69,0.5)' } },
            }
        };
    }

    function destroyChart(k) { if (chartInstances[k]) { chartInstances[k].destroy(); delete chartInstances[k]; } }

    function renderValConvByBranch() {
        destroyChart('valConvBranch');
        const pG = groupBy(filteredProduct, 'branch');
        const oG = groupBy(filteredOSG, 'branch');
        const entries = Object.keys(pG).map(name => {
            const pSold = pG[name].reduce((s, r) => s + r.soldPrice, 0);
            const oSold = (oG[name] || []).reduce((s, r) => s + r.soldPrice, 0);
            return { name, conv: pSold > 0 ? (oSold / pSold) * 100 : 0 };
        }).sort((a, b) => b.conv - a.conv).slice(0, 15);

        chartInstances['valConvBranch'] = new Chart($('chartValConvBranch').getContext('2d'), {
            type: 'bar',
            data: {
                labels: entries.map(e => truncate(e.name, 16)),
                datasets: [{ label: 'Value Conv %', data: entries.map(e => e.conv), backgroundColor: COLORS.amber, borderRadius: 4 }]
            },
            options: { ...chartBase(false), plugins: { ...chartBase(false).plugins, legend: { display: false } } }
        });
    }

    function renderQtyConvByRBM() {
        destroyChart('qtyConvRBM');
        const pG = groupBy(filteredProduct, 'rbm');

        // Map OSG → RBM via invoice lookup (same as renderConvTable)
        const invoiceToRBM = {};
        filteredProduct.forEach(r => { if (r.invoice && r.rbm) invoiceToRBM[r.invoice] = r.rbm; });
        const osgByRBM = {};
        filteredOSG.forEach(r => {
            const rbm = r.invoice ? (invoiceToRBM[r.invoice] || null) : null;
            if (!rbm) return; // skip unattributable rows — no "Unknown"
            if (!osgByRBM[rbm]) osgByRBM[rbm] = [];
            osgByRBM[rbm].push(r);
        });

        const entries = Object.keys(pG).map(name => {
            const pQty = pG[name].reduce((s, r) => s + r.qty, 0);
            const oQty = (osgByRBM[name] || []).reduce((s, r) => s + r.qty, 0);
            return { name, conv: pQty > 0 ? (oQty / pQty) * 100 : 0 };
        }).sort((a, b) => b.conv - a.conv).slice(0, 10);

        chartInstances['qtyConvRBM'] = new Chart($('chartQtyConvRBM').getContext('2d'), {
            type: 'bar',
            data: {
                labels: entries.map(e => e.name),
                datasets: [{ label: 'Qty Conv %', data: entries.map(e => e.conv), backgroundColor: COLORS.orange, borderRadius: 4 }]
            },
            options: { ...chartBase(false), indexAxis: 'y', plugins: { ...chartBase(false).plugins, legend: { display: false } } }
        });
    }

    function renderQtyConvByProduct() {
        destroyChart('qtyConvProduct');
        const pG = groupBy(filteredProduct, 'product');
        const oG = groupBy(filteredOSG, 'product');
        const entries = Object.keys(pG).map(name => {
            const pQty = pG[name].reduce((s, r) => s + r.qty, 0);
            const oQty = (oG[name] || []).reduce((s, r) => s + r.qty, 0);
            return { name, conv: pQty > 0 ? (oQty / pQty) * 100 : 0 };
        }).sort((a, b) => b.conv - a.conv).slice(0, 10);

        chartInstances['qtyConvProduct'] = new Chart($('chartQtyConvProduct').getContext('2d'), {
            type: 'bar',
            data: {
                labels: entries.map(e => truncate(e.name, 14)),
                datasets: [{ label: 'Qty Conv %', data: entries.map(e => e.conv), backgroundColor: PALETTE.slice(0, entries.length), borderRadius: 4 }]
            },
            options: { ...chartBase(false), plugins: { ...chartBase(false).plugins, legend: { display: false } } }
        });
    }


    function renderProductRevenueChart() {
        destroyChart('productRev');
        const grouped = groupBy(filteredAll, 'product');
        const entries = Object.entries(grouped)
            .map(([name, rows]) => ({ name, rev: rows.reduce((s, r) => s + r.soldPrice, 0) }))
            .sort((a, b) => b.rev - a.rev).slice(0, 10);

        chartInstances['productRev'] = new Chart($('chartProductRevenue').getContext('2d'), {
            type: 'bar',
            data: {
                labels: entries.map(e => truncate(e.name, 14)),
                datasets: [{ label: 'Sold Price Total', data: entries.map(e => e.rev), backgroundColor: PALETTE.slice(0, entries.length), borderRadius: 4 }]
            },
            options: { ...chartBase(true), plugins: { ...chartBase(true).plugins, legend: { display: false } } }
        });
    }

    function renderProductPieChart() {
        destroyChart('productPie');
        const grouped = groupBy(filteredAll, 'product');
        const entries = Object.entries(grouped)
            .map(([name, rows]) => ({ name, rev: rows.reduce((s, r) => s + r.soldPrice, 0) }))
            .sort((a, b) => b.rev - a.rev).slice(0, 8);

        chartInstances['productPie'] = new Chart($('chartProductPie').getContext('2d'), {
            type: 'doughnut',
            data: {
                labels: entries.map(e => e.name),
                datasets: [{ data: entries.map(e => e.rev), backgroundColor: PALETTE.slice(0, entries.length), borderWidth: 0 }]
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                plugins: { legend: { position: 'right', labels: { color: '#8896b8', font: { family: 'Inter', size: 11 }, padding: 12 } } }
            }
        });
    }

    // ---- EXPORT CSV ----
    document.querySelectorAll('.btn-export').forEach(btn => {
        btn.addEventListener('click', () => {
            const type = btn.dataset.export;
            if (type === 'conversion') { exportConversionCSV(); return; }
            let rows = filteredAll;
            exportCSV(rows, type + '_report.csv');
        });
    });

    function exportCSV(rows, filename) {
        if (rows.length === 0) return;
        const hdr = ['Branch', 'RBM', 'BDM', 'Staff', 'Product', 'Brand', 'Sold Price', 'Taxable Value', 'Tax', 'QTY', 'Profit/Loss'];
        const lines = [hdr.join(',')];
        rows.forEach(r => {
            lines.push([q(r.branch), q(r.rbm), q(r.bdm), q(r.staff), q(r.product), q(r.brand),
            r.soldPrice, r.taxableVal, r.tax, r.qty, r.profit.toFixed(2)].join(','));
        });
        downloadCSV(lines.join('\n'), filename);
    }

    function exportConversionCSV() {
        const pG = groupBy(filteredProduct, 'branch');
        const oG = groupBy(filteredOSG, 'branch');
        const allKeys = new Set([...Object.keys(pG), ...Object.keys(oG)]);
        const hdr = ['Branch', 'Product Revenue', 'OSG Revenue', 'Value Conv%', 'Product Qty', 'OSG Qty', 'Qty Conv%'];
        const lines = [hdr.join(',')];
        allKeys.forEach(branch => {
            const pRows = pG[branch] || []; const oRows = oG[branch] || [];
            const pS = pRows.reduce((s, r) => s + r.soldPrice, 0);
            const oS = oRows.reduce((s, r) => s + r.soldPrice, 0);
            const pQ = pRows.reduce((s, r) => s + r.qty, 0);
            const oQ = oRows.reduce((s, r) => s + r.qty, 0);
            const vc = pS > 0 ? ((oS / pS) * 100).toFixed(2) : '0';
            const qc = pQ > 0 ? ((oQ / pQ) * 100).toFixed(2) : '0';
            lines.push([q(branch), pS, oS, vc, pQ, oQ, qc].join(','));
        });
        downloadCSV(lines.join('\n'), 'conversion_report.csv');
    }

    function downloadCSV(content, filename) {
        const blob = new Blob([content], { type: 'text/csv' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a'); a.href = url; a.download = filename; a.click();
        URL.revokeObjectURL(url);
    }
    function q(s) { return '"' + String(s).replace(/"/g, '""') + '"'; }

    // ---- RESET ----
    btnReset.addEventListener('click', () => {
        productData = []; osgData = []; amcData = []; allData = [];
        filteredProduct = []; filteredOSG = []; filteredAll = [];
        productStatus.className = 'upload-status'; productStatus.innerHTML = '';
        osgStatus.className = 'upload-status'; osgStatus.innerHTML = '';
        amcStatus.className = 'upload-status'; amcStatus.innerHTML = '';
        fileCountBadge.style.display = 'none'; btnReset.style.display = 'none';
        btnGenerate.disabled = true;
        [filterRBM, filterBranch, filterBDM, filterStaff].forEach(sel => {
            sel.innerHTML = `<option value="">${sel.options[0]?.textContent || 'All'}</option>`;
        });
        // Product has hardcoded options — just reset selection
        filterProduct.value = '';
        // Brand has dynamic options — clear and reset
        filterBrand.innerHTML = '<option value="">All Brands</option>';
        $('kpiValConv').textContent = '0%';
        $('kpiQtyConv').textContent = '0%'; $('kpiQuantity').textContent = '0';
        ['convRBMTable', 'convBDMTable', 'convStaffTable', 'convBranchTable', 'convProductTable'].forEach(id => $(id).innerHTML = '');
        ['conversionTableWrapper', 'fullDataTableWrapper'].forEach(id => $(id).innerHTML = '');
        Object.keys(chartInstances).forEach(destroyChart);
        $('lcTableWrapper').innerHTML = noDataHTML('Upload data and generate reports first.');
        $('lcKpiRow').innerHTML = '';
        $('lcCount').textContent = '0 staff';
        $('tcTableWrapper').innerHTML = noDataHTML('Upload data and generate reports first.');
        $('tcKpiRow').innerHTML = '';
        $('tcCount').textContent = '0 staff';
        $('insightsContent').innerHTML = noDataHTML('Upload data and generate reports to see insights.');
        document.querySelector('[data-section="upload-section"]').click();
    });

    // ---- LOW CONV STAFF PAGE ----
    $('btnLCRefresh').addEventListener('click', renderLowConvPage);
    $('btnLCExport').addEventListener('click', exportLowConvCSV);

    // Also refresh whenever user navigates to the page
    document.querySelector('[data-section="lowconv-section"]').addEventListener('click', () => {
        setTimeout(renderLowConvPage, 50);
    });

    // ---- LOW CONV STAFF LOGIC ----
    function buildStaffStats() {
        // Build invoice → staff lookup from ALL product data (unfiltered)
        const invoiceStaff = {};
        productData.forEach(r => { if (r.invoice && r.staff) invoiceStaff[r.invoice] = r.staff; });

        // Group product data by staff
        const pByStaff = {};
        productData.forEach(r => {
            const s = r.staff || 'Unknown';
            if (!pByStaff[s]) pByStaff[s] = { branch: r.branch, rbm: r.rbm, bdm: r.bdm, rows: [] };
            pByStaff[s].rows.push(r);
        });

        // Group OSG data by staff via invoice mapping
        const oByStaff = {};
        osgData.forEach(r => {
            const s = r.invoice ? (invoiceStaff[r.invoice] || null) : null;
            if (!s) return;
            if (!oByStaff[s]) oByStaff[s] = [];
            oByStaff[s].push(r);
        });

        const allStaff = new Set([...Object.keys(pByStaff), ...Object.keys(oByStaff)]);
        allStaff.delete('Unknown');

        return Array.from(allStaff).map(name => {
            const pInfo = pByStaff[name] || { branch: '', rbm: '', bdm: '', rows: [] };
            const oRows = oByStaff[name] || [];
            const pQty = pInfo.rows.reduce((s, r) => s + r.qty, 0);
            const oQty = oRows.reduce((s, r) => s + r.qty, 0);
            const pRev = pInfo.rows.reduce((s, r) => s + r.soldPrice, 0);
            const oRev = oRows.reduce((s, r) => s + r.soldPrice, 0);
            const qtyConv = pQty > 0 ? (oQty / pQty) * 100 : 0;
            const valConv = pRev > 0 ? (oRev / pRev) * 100 : 0;
            return { name, branch: pInfo.branch, rbm: pInfo.rbm, bdm: pInfo.bdm, pQty, oQty, pRev, oRev, qtyConv, valConv };
        });
    }

    function renderLowConvPage() {
        if (productData.length === 0) {
            $('lcTableWrapper').innerHTML = noDataHTML('Upload data and generate reports first.');
            $('lcKpiRow').innerHTML = '';
            $('lcCount').textContent = '0 staff';
            return;
        }

        const minQty = parseFloat($('lcMinQty').value) || 0;
        const maxConv = parseFloat($('lcMaxConv').value);

        const allStats = buildStaffStats();

        // Filter: must have >= minQty product qty AND <= maxConv qty conversion %
        const filtered = allStats
            .filter(s => s.pQty >= minQty && s.qtyConv <= maxConv)
            .sort((a, b) => {
                // Primary: lowest qty conversion first
                if (a.qtyConv !== b.qtyConv) return a.qtyConv - b.qtyConv;
                // Secondary: highest product qty first
                return b.pQty - a.pQty;
            });

        $('lcCount').textContent = `${filtered.length} staff`;

        // KPI summary
        const totalPQty = filtered.reduce((s, r) => s + r.pQty, 0);
        const totalOQty = filtered.reduce((s, r) => s + r.oQty, 0);
        const totalPRev = filtered.reduce((s, r) => s + r.pRev, 0);
        const zeroConvCount = filtered.filter(r => r.qtyConv === 0).length;
        $('lcKpiRow').innerHTML = `
            <div class="lc-kpi"><span class="lc-kpi-label">Zero Conv Staff</span><span class="lc-kpi-val loss-text">${zeroConvCount}</span></div>
            <div class="lc-kpi"><span class="lc-kpi-label">Total Product Qty</span><span class="lc-kpi-val">${formatNumber(totalPQty)}</span></div>
            <div class="lc-kpi"><span class="lc-kpi-label">Total OSG Qty (Sold)</span><span class="lc-kpi-val">${formatNumber(totalOQty)}</span></div>
            <div class="lc-kpi"><span class="lc-kpi-label">Opportunity Missed (Qty)</span><span class="lc-kpi-val loss-text">${formatNumber(totalPQty - totalOQty)}</span></div>
            <div class="lc-kpi"><span class="lc-kpi-label">Total Product Revenue</span><span class="lc-kpi-val">${fmtShort(totalPRev)}</span></div>
        `;

        if (filtered.length === 0) {
            $('lcTableWrapper').innerHTML = noDataHTML(`No staff found with ≥${minQty} product qty and ≤${maxConv}% qty conversion.`);
            return;
        }

        let html = `<table class="data-table">
            <thead><tr>
                <th>#</th><th>Staff</th><th>Branch</th><th>RBM</th><th>BDM</th>
                <th>Prod Qty</th><th>OSG Qty</th><th>Qty Conv%</th><th>Val Conv%</th><th>Prod Rev</th>
            </tr></thead><tbody>`;

        filtered.forEach((e, i) => {
            const convCls = e.qtyConv === 0 ? 'loss-val' : (e.qtyConv < 0.5 ? 'conv-warn' : 'conv-val');
            const rank = i + 1;
            const rankBadge = rank <= 3 ? `<span class="rank-badge rank-${rank}">${rank}</span>` : `<span class="rank-num">${rank}</span>`;
            html += `<tr>
                <td class="number-cell">${rankBadge}</td>
                <td><strong>${e.name}</strong></td>
                <td>${e.branch}</td>
                <td>${e.rbm}</td>
                <td>${e.bdm}</td>
                <td class="number-cell"><strong>${e.pQty}</strong></td>
                <td class="number-cell">${e.oQty}</td>
                <td class="number-cell ${convCls}">${e.qtyConv.toFixed(2)}%</td>
                <td class="number-cell conv-val">${e.valConv.toFixed(2)}%</td>
                <td class="number-cell">${fmtShort(e.pRev)}</td>
            </tr>`;
        });

        html += '</tbody></table>';
        $('lcTableWrapper').innerHTML = html;
    }

    function exportLowConvCSV() {
        if (productData.length === 0) return;
        const minQty = parseFloat($('lcMinQty').value) || 0;
        const maxConv = parseFloat($('lcMaxConv').value);
        const allStats = buildStaffStats();
        const filtered = allStats.filter(s => s.pQty >= minQty && s.qtyConv <= maxConv)
            .sort((a, b) => a.qtyConv - b.qtyConv || b.pQty - a.pQty);
        if (filtered.length === 0) return;
        const hdr = ['Rank', 'Staff', 'Branch', 'RBM', 'BDM', 'Prod Qty', 'OSG Qty', 'Qty Conv%', 'Val Conv%', 'Prod Revenue'];
        const lines = [hdr.join(',')];
        filtered.forEach((e, i) => {
            lines.push([i + 1, q(e.name), q(e.branch), q(e.rbm), q(e.bdm), e.pQty, e.oQty,
            e.qtyConv.toFixed(2), e.valConv.toFixed(2), e.pRev.toFixed(0)].join(','));
        });
        downloadCSV(lines.join('\n'), 'low_conv_staff.csv');
    }

    // ---- TOP CONV STAFF PAGE ----
    $('btnTCRefresh').addEventListener('click', renderTopConvPage);
    $('btnTCExport').addEventListener('click', exportTopConvCSV);
    document.querySelector('[data-section="topconv-section"]').addEventListener('click', () => {
        setTimeout(renderTopConvPage, 50);
    });

    function renderTopConvPage() {
        if (productData.length === 0) {
            $('tcTableWrapper').innerHTML = noDataHTML('Upload data and generate reports first.');
            $('tcKpiRow').innerHTML = '';
            $('tcCount').textContent = '0 staff';
            return;
        }

        const minQty = parseFloat($('tcMinQty').value) || 0;
        const sortBy = $('tcSortBy').value; // 'qtyConv' or 'valConv'
        const topN = parseInt($('tcTopN').value) || 50;

        const allStats = buildStaffStats();

        // Filter: must have >= minQty product qty AND conversion > 0
        const filtered = allStats
            .filter(s => s.pQty >= minQty && s[sortBy] > 0)
            .sort((a, b) => b[sortBy] - a[sortBy])
            .slice(0, topN);

        $('tcCount').textContent = `${filtered.length} staff`;

        // KPI summary
        const avgQtyConv = filtered.length > 0 ? filtered.reduce((s, r) => s + r.qtyConv, 0) / filtered.length : 0;
        const avgValConv = filtered.length > 0 ? filtered.reduce((s, r) => s + r.valConv, 0) / filtered.length : 0;
        const totalOQty = filtered.reduce((s, r) => s + r.oQty, 0);
        const totalORev = filtered.reduce((s, r) => s + r.oRev, 0);
        const totalPRev = filtered.reduce((s, r) => s + r.pRev, 0);
        $('tcKpiRow').innerHTML = `
            <div class="lc-kpi"><span class="lc-kpi-label">Avg Qty Conv</span><span class="lc-kpi-val profit-text">${avgQtyConv.toFixed(2)}%</span></div>
            <div class="lc-kpi"><span class="lc-kpi-label">Avg Val Conv</span><span class="lc-kpi-val profit-text">${avgValConv.toFixed(2)}%</span></div>
            <div class="lc-kpi"><span class="lc-kpi-label">Total OSG Qty</span><span class="lc-kpi-val">${formatNumber(totalOQty)}</span></div>
            <div class="lc-kpi"><span class="lc-kpi-label">Total OSG Revenue</span><span class="lc-kpi-val">${fmtShort(totalORev)}</span></div>
            <div class="lc-kpi"><span class="lc-kpi-label">Total Prod Revenue</span><span class="lc-kpi-val">${fmtShort(totalPRev)}</span></div>
        `;

        if (filtered.length === 0) {
            $('tcTableWrapper').innerHTML = noDataHTML('No staff found matching criteria.');
            return;
        }

        const sortLabel = sortBy === 'qtyConv' ? 'Qty Conv%' : 'Val Conv%';
        let html = `<table class="data-table">
            <thead><tr>
                <th>#</th><th>Staff</th><th>Branch</th><th>RBM</th><th>BDM</th>
                <th>Prod Qty</th><th>OSG Qty</th><th>Qty Conv%</th><th>Val Conv%</th><th>Prod Rev</th><th>OSG Rev</th>
            </tr></thead><tbody>`;

        filtered.forEach((e, i) => {
            const rank = i + 1;
            const rankBadge = rank <= 3 ? `<span class="rank-badge rank-${rank}">${rank}</span>` : `<span class="rank-num">${rank}</span>`;
            html += `<tr>
                <td class="number-cell">${rankBadge}</td>
                <td><strong>${e.name}</strong></td>
                <td>${e.branch}</td>
                <td>${e.rbm}</td>
                <td>${e.bdm}</td>
                <td class="number-cell">${e.pQty}</td>
                <td class="number-cell"><strong>${e.oQty}</strong></td>
                <td class="number-cell profit-val">${e.qtyConv.toFixed(2)}%</td>
                <td class="number-cell profit-val">${e.valConv.toFixed(2)}%</td>
                <td class="number-cell">${fmtShort(e.pRev)}</td>
                <td class="number-cell">${fmtShort(e.oRev)}</td>
            </tr>`;
        });

        html += '</tbody></table>';
        $('tcTableWrapper').innerHTML = html;
    }

    function exportTopConvCSV() {
        if (productData.length === 0) return;
        const minQty = parseFloat($('tcMinQty').value) || 0;
        const sortBy = $('tcSortBy').value;
        const topN = parseInt($('tcTopN').value) || 50;
        const allStats = buildStaffStats();
        const filtered = allStats.filter(s => s.pQty >= minQty && s[sortBy] > 0)
            .sort((a, b) => b[sortBy] - a[sortBy]).slice(0, topN);
        if (filtered.length === 0) return;
        const hdr = ['Rank', 'Staff', 'Branch', 'RBM', 'BDM', 'Prod Qty', 'OSG Qty', 'Qty Conv%', 'Val Conv%', 'Prod Revenue', 'OSG Revenue'];
        const lines = [hdr.join(',')];
        filtered.forEach((e, i) => {
            lines.push([i + 1, q(e.name), q(e.branch), q(e.rbm), q(e.bdm), e.pQty, e.oQty,
            e.qtyConv.toFixed(2), e.valConv.toFixed(2), e.pRev.toFixed(0), e.oRev.toFixed(0)].join(','));
        });
        downloadCSV(lines.join('\n'), 'top_conv_staff.csv');
    }

    // ---- DEEP INSIGHTS PAGE ----
    $('btnInsightsRefresh').addEventListener('click', renderInsightsPage);
    document.querySelector('[data-section="insights-section"]').addEventListener('click', () => {
        setTimeout(renderInsightsPage, 50);
    });

    function renderInsightsPage() {
        if (productData.length === 0) {
            $('insightsContent').innerHTML = noDataHTML('Upload data and generate reports to see insights.');
            return;
        }

        const staffStats = buildStaffStats();
        const conv = calcConversion(productData, osgData);
        const totalStaff = staffStats.length;
        const zeroConvStaff = staffStats.filter(s => s.qtyConv === 0 && s.pQty > 0);
        const topQty = [...staffStats].filter(s => s.pQty >= 3).sort((a, b) => b.qtyConv - a.qtyConv).slice(0, 5);
        const topVal = [...staffStats].filter(s => s.pQty >= 3).sort((a, b) => b.valConv - a.valConv).slice(0, 5);

        // Branch stats
        const pByBranch = groupBy(productData, 'branch');
        const invoiceBranch = {};
        productData.forEach(r => { if (r.invoice && r.branch) invoiceBranch[r.invoice] = r.branch; });
        const oByBranch = {};
        osgData.forEach(r => {
            const b = r.branch || (r.invoice ? invoiceBranch[r.invoice] : null) || 'Unknown';
            if (!oByBranch[b]) oByBranch[b] = [];
            oByBranch[b].push(r);
        });
        const branchStats = Object.keys(pByBranch).map(b => {
            const pRows = pByBranch[b] || [];
            const oRows = oByBranch[b] || [];
            const pQty = pRows.reduce((s, r) => s + r.qty, 0);
            const oQty = oRows.reduce((s, r) => s + r.qty, 0);
            const pRev = pRows.reduce((s, r) => s + r.soldPrice, 0);
            const oRev = oRows.reduce((s, r) => s + r.soldPrice, 0);
            return { name: b, pQty, oQty, pRev, oRev, qtyConv: pQty > 0 ? (oQty / pQty) * 100 : 0, valConv: pRev > 0 ? (oRev / pRev) * 100 : 0 };
        });

        // RBM stats
        const pByRBM = groupBy(productData, 'rbm');
        const invoiceRBM = {};
        productData.forEach(r => { if (r.invoice && r.rbm) invoiceRBM[r.invoice] = r.rbm; });
        const oByRBM = {};
        osgData.forEach(r => {
            const rbm = r.invoice ? (invoiceRBM[r.invoice] || null) : null;
            if (rbm) { if (!oByRBM[rbm]) oByRBM[rbm] = []; oByRBM[rbm].push(r); }
        });
        const rbmStats = Object.keys(pByRBM).filter(k => k !== 'Unknown').map(name => {
            const pRows = pByRBM[name] || [];
            const oRows = oByRBM[name] || [];
            const pQty = pRows.reduce((s, r) => s + r.qty, 0);
            const oQty = oRows.reduce((s, r) => s + r.qty, 0);
            const pRev = pRows.reduce((s, r) => s + r.soldPrice, 0);
            return { name, pQty, oQty, pRev, qtyConv: pQty > 0 ? (oQty / pQty) * 100 : 0 };
        });

        // Product stats
        const pByProd = groupBy(productData, 'product');
        const oByProd = groupBy(osgData, 'product');
        const prodStats = Object.keys(pByProd).map(name => {
            const pRows = pByProd[name] || [];
            const oRows = oByProd[name] || [];
            const pQty = pRows.reduce((s, r) => s + r.qty, 0);
            const oQty = oRows.reduce((s, r) => s + r.qty, 0);
            const pRev = pRows.reduce((s, r) => s + r.soldPrice, 0);
            return { name, pQty, oQty, pRev, qtyConv: pQty > 0 ? (oQty / pQty) * 100 : 0 };
        });

        // Revenue concentration
        const totalPRev = productData.reduce((s, r) => s + r.soldPrice, 0);
        const branchRevShare = branchStats.map(b => ({ name: b.name, share: totalPRev > 0 ? (b.pRev / totalPRev) * 100 : 0 })).sort((a, b) => b.share - a.share);

        let html = '';

        // ---- Card 1: Overall Summary ----
        html += insightCard('📊', 'Overall Performance Summary', 'info', `
            <div class="insight-metrics">
                <div class="insight-metric"><span class="metric-val">${formatNumber(productData.length)}</span><span class="metric-label">Total Transactions</span></div>
                <div class="insight-metric"><span class="metric-val">${totalStaff}</span><span class="metric-label">Active Staff</span></div>
                <div class="insight-metric"><span class="metric-val">${conv.valueConv.toFixed(2)}%</span><span class="metric-label">Value Conversion</span></div>
                <div class="insight-metric"><span class="metric-val">${conv.qtyConv.toFixed(2)}%</span><span class="metric-label">Qty Conversion</span></div>
                <div class="insight-metric"><span class="metric-val">${fmtShort(totalPRev)}</span><span class="metric-label">Total Prod Revenue</span></div>
                <div class="insight-metric"><span class="metric-val">${Object.keys(pByBranch).length}</span><span class="metric-label">Active Branches</span></div>
            </div>
        `);

        // ---- Card 2: Zero Conversion Alert ----
        if (zeroConvStaff.length > 0) {
            const zeroTotalQty = zeroConvStaff.reduce((s, r) => s + r.pQty, 0);
            const zeroTotalRev = zeroConvStaff.reduce((s, r) => s + r.pRev, 0);
            const topZero = zeroConvStaff.sort((a, b) => b.pQty - a.pQty).slice(0, 5);
            html += insightCard('🚨', `Zero Conversion Alert — ${zeroConvStaff.length} Staff`, 'danger', `
                <p><strong>${zeroConvStaff.length} staff</strong> have sold <strong>${formatNumber(zeroTotalQty)} products</strong> (${fmtShort(zeroTotalRev)} revenue) but <strong>zero OSG/warranty conversion</strong>.</p>
                <div class="insight-tag-row">
                    ${topZero.map(s => `<span class="insight-tag danger">${s.name} (${s.pQty} qty)</span>`).join('')}
                    ${zeroConvStaff.length > 5 ? `<span class="insight-tag muted">+${zeroConvStaff.length - 5} more</span>` : ''}
                </div>
                <div class="insight-solution">
                    <strong>💡 Solution:</strong> Conduct targeted training for these staff members on OSG selling techniques. Pair them with top converters for mentorship. Set 1-week conversion targets with incentives.
                </div>
            `);
        }

        // ---- Card 3: Top Performers ----
        if (topQty.length > 0) {
            html += insightCard('🏆', 'Top Performers — Best Qty Conversion', 'success', `
                <div class="insight-list">
                    ${topQty.map((s, i) => `
                        <div class="insight-list-item">
                            <span class="rank-badge rank-${i < 3 ? i + 1 : 'n'}">${i + 1}</span>
                            <strong>${s.name}</strong>
                            <span class="insight-pill success">${s.qtyConv.toFixed(1)}% qty conv</span>
                            <span class="insight-pill info">${s.pQty} products</span>
                            <span class="text-muted">${s.branch}</span>
                        </div>
                    `).join('')}
                </div>
                <div class="insight-solution">
                    <strong>💡 Recommendation:</strong> Recognize these staff publicly. Study their techniques and replicate across other branches. Consider a reward/incentive program to sustain performance.
                </div>
            `);
        }

        // ---- Card 4: Underperforming Branches ----
        const weakBranches = branchStats.filter(b => b.pQty >= 10 && b.qtyConv < 2).sort((a, b) => a.qtyConv - b.qtyConv).slice(0, 5);
        if (weakBranches.length > 0) {
            html += insightCard('📉', 'Underperforming Branches', 'warning', `
                <p>These branches have significant product sales but very low OSG conversion:</p>
                <table class="data-table insight-table"><thead><tr>
                    <th>Branch</th><th>Prod Qty</th><th>OSG Qty</th><th>Qty Conv%</th>
                </tr></thead><tbody>
                    ${weakBranches.map(b => `<tr><td>${b.name}</td><td class="number-cell">${b.pQty}</td><td class="number-cell">${b.oQty}</td><td class="number-cell loss-val">${b.qtyConv.toFixed(2)}%</td></tr>`).join('')}
                </tbody></table>
                <div class="insight-solution">
                    <strong>💡 Solution:</strong> Schedule branch visits and OSG training workshops. Review branch-level OSG targets. Investigate if product mix or customer demographics contribute to low conversion.
                </div>
            `);
        }

        // ---- Card 5: RBM Performance Gap ----
        if (rbmStats.length >= 2) {
            const rbmSorted = [...rbmStats].sort((a, b) => b.qtyConv - a.qtyConv);
            const best = rbmSorted[0];
            const worst = rbmSorted[rbmSorted.length - 1];
            const gap = best.qtyConv - worst.qtyConv;
            html += insightCard('👥', 'RBM Performance Gap', gap > 5 ? 'warning' : 'info', `
                <div class="insight-compare">
                    <div class="compare-box success-bg">
                        <span class="compare-label">Best RBM</span>
                        <strong>${best.name}</strong>
                        <span class="insight-pill success">${best.qtyConv.toFixed(2)}% qty conv</span>
                    </div>
                    <div class="compare-divider">vs</div>
                    <div class="compare-box danger-bg">
                        <span class="compare-label">Needs Improvement</span>
                        <strong>${worst.name}</strong>
                        <span class="insight-pill danger">${worst.qtyConv.toFixed(2)}% qty conv</span>
                    </div>
                </div>
                <p>Performance gap: <strong>${gap.toFixed(2)}%</strong>. ${gap > 5 ? 'This is a significant gap that needs attention.' : 'Relatively close performance.'}</p>
                <div class="insight-solution">
                    <strong>💡 Recommendation:</strong> ${gap > 5 ? 'Organize knowledge-sharing sessions between top and bottom RBMs. Assign mentors and set improvement timelines.' : 'Performance is fairly balanced. Focus on pushing overall numbers higher.'}
                </div>
            `);
        }

        // ---- Card 6: Product Category Analysis ----
        const prodSorted = [...prodStats].filter(p => p.pQty >= 5).sort((a, b) => a.qtyConv - b.qtyConv);
        const weakProds = prodSorted.slice(0, 3);
        const strongProds = prodSorted.slice(-3).reverse();
        if (prodSorted.length > 0) {
            html += insightCard('📦', 'Product Category Analysis', 'info', `
                <div class="insight-compare">
                    <div class="compare-box success-bg" style="flex:1;">
                        <span class="compare-label">Strong Categories</span>
                        ${strongProds.map(p => `<div class="mini-row"><strong>${p.name}</strong> <span class="insight-pill success">${p.qtyConv.toFixed(1)}%</span></div>`).join('')}
                    </div>
                    <div class="compare-box danger-bg" style="flex:1;">
                        <span class="compare-label">Weak Categories</span>
                        ${weakProds.map(p => `<div class="mini-row"><strong>${p.name}</strong> <span class="insight-pill danger">${p.qtyConv.toFixed(1)}%</span></div>`).join('')}
                    </div>
                </div>
                <div class="insight-solution">
                    <strong>💡 Solution:</strong> Focus OSG push on weak categories. Create category-specific sales scripts. Consider bundled OSG offers for low-converting product types.
                </div>
            `);
        }

        // ---- Card 7: Revenue Concentration Risk ----
        if (branchRevShare.length >= 3) {
            const top3Share = branchRevShare.slice(0, 3).reduce((s, b) => s + b.share, 0);
            html += insightCard('⚖️', 'Revenue Concentration', top3Share > 50 ? 'warning' : 'info', `
                <p>Top 3 branches contribute <strong>${top3Share.toFixed(1)}%</strong> of total product revenue:</p>
                <div class="insight-tag-row">
                    ${branchRevShare.slice(0, 5).map(b => `<span class="insight-tag info">${b.name}: ${b.share.toFixed(1)}%</span>`).join('')}
                </div>
                ${top3Share > 50 ? '<p class="text-warning">⚠️ High concentration risk — underperformance in these branches would significantly impact overall numbers.</p>' : '<p class="text-success">✅ Revenue is fairly distributed — good diversification.</p>'}
                <div class="insight-solution">
                    <strong>💡 Recommendation:</strong> ${top3Share > 50 ? 'Invest in growing smaller branches. Reduce dependency on top branches by improving performance of bottom 50%.' : 'Maintain balanced growth across all branches.'}
                </div>
            `);
        }

        // ---- Card 8: Action Plan ----
        const urgentActions = [];
        if (zeroConvStaff.length > 5) urgentActions.push(`Train ${zeroConvStaff.length} zero-conversion staff on OSG selling immediately`);
        if (weakBranches.length > 0) urgentActions.push(`Conduct branch visits to ${weakBranches.map(b => b.name).join(', ')}`);
        if (conv.qtyConv < 5) urgentActions.push(`Overall qty conversion (${conv.qtyConv.toFixed(1)}%) is below target — launch org-wide OSG campaign`);
        urgentActions.push('Review and update staff-wise weekly conversion targets');
        urgentActions.push('Share top performer success stories in team meetings');
        if (topQty.length > 0) urgentActions.push(`Reward top converters: ${topQty.slice(0, 3).map(s => s.name).join(', ')}`);

        html += insightCard('🎯', 'Action Plan — Next Steps', 'action', `
            <ol class="insight-actions">
                ${urgentActions.map(a => `<li>${a}</li>`).join('')}
            </ol>
        `);

        $('insightsContent').innerHTML = html;
    }

    function insightCard(icon, title, severity, body) {
        return `<div class="insight-card insight-${severity}">
            <div class="insight-card-header">
                <span class="insight-icon">${icon}</span>
                <h3>${title}</h3>
            </div>
            <div class="insight-card-body">${body}</div>
        </div>`;
    }

    // ---- UTILITIES ----
    function groupBy(arr, key) {
        const m = {};
        arr.forEach(r => { const k = r[key] || 'Unknown'; if (!m[k]) m[k] = []; m[k].push(r); });
        return m;
    }
    function formatCurrency(n) { return '₹' + n.toLocaleString('en-IN', { maximumFractionDigits: 0 }); }
    function fmtShort(n) {
        if (Math.abs(n) >= 1e7) return '₹' + (n / 1e7).toFixed(1) + 'Cr';
        if (Math.abs(n) >= 1e5) return '₹' + (n / 1e5).toFixed(1) + 'L';
        if (Math.abs(n) >= 1e3) return '₹' + (n / 1e3).toFixed(1) + 'K';
        return '₹' + n.toFixed(0);
    }
    function formatNumber(n) { return n.toLocaleString('en-IN'); }
    function capitalize(s) { return s.charAt(0).toUpperCase() + s.slice(1); }
    function truncate(s, len) { return s.length > len ? s.substring(0, len) + '…' : s; }
    function showLoading(show) { loadingOverlay.classList.toggle('active', show); }

})();
