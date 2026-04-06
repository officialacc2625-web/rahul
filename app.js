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

    // ---- FIREBASE INIT & SHARE LINK ----
    const firebaseConfig = {
        apiKey: "AIzaSyC8QKzFwgYnONyp8182KjV_SGuiD6cGPUc",
        authDomain: "myg-analytics.firebaseapp.com",
        projectId: "myg-analytics",
        storageBucket: "myg-analytics.firebasestorage.app",
        messagingSenderId: "126534817953",
        appId: "1:126534817953:web:300aefd1de65da8242f8e0",
        databaseURL: "https://myg-analytics-default-rtdb.firebaseio.com" // Fallback if missing
    };
    if (typeof firebase !== 'undefined' && !firebase.apps.length) {
        firebase.initializeApp(firebaseConfig);
    }

    const shareParam = new URLSearchParams(window.location.search).get('share');
    if (shareParam) {
        window.isSharedView = true;
        document.addEventListener('DOMContentLoaded', () => {
            const loadingOverlay = document.getElementById('loadingOverlay');
            if (loadingOverlay) loadingOverlay.style.display = 'flex';

            firebase.database().ref('shares/' + shareParam).once('value').then(snap => {
                if (loadingOverlay) loadingOverlay.style.display = 'none';
                const data = snap.val();
                if (data) {
                    window.sharedMissedUnique = data.missedUnique || [];
                    // Do NOT set isAuthenticated=true — keep all other pages locked
                    document.querySelector('[data-section="customers-osg-section"]').click();
                } else {
                    alert('Share link is invalid or expired.');
                }
            }).catch(e => {
                if (loadingOverlay) loadingOverlay.style.display = 'none';
                alert('Failed to load shared link.');
            });
        });
    }

    // ---- COLUMN MAPPING (Product / AMC file) ----
    const PRODUCT_COL_MAP = {
        branch: ['branch', 'store name', 'store', 'branch name', 'outlet', 'outlet name', 'shop name'],
        rbm: ['rbm', 'rbm name', 'region', 'regional manager', 'regional business manager', 'rsm'],
        bdm: ['bdm', 'bdm name', 'business development manager', 'area manager', 'asm'],
        staff: ['staff', 'staff name', 'salesperson', 'sales person', 'employee', 'employee name', 'promoter', 'promoter name', 'executive'],
        product: ['product', 'product name', 'product type', 'product group', 'model', 'model name', 'item name'],
        category: ['category', 'item category', 'item group', 'product category', 'sub category'],
        brand: ['brand', 'brand name', 'make'],
        soldPrice: ['sold price', 'soldprice', 'selling price', 'sale price', 'mop', 'net amount', 'net value', 'total amount', 'amount', 'sale amount', 'sale value', 'value', 'net sales value'],
        taxableVal: ['taxable value', 'taxable', 'taxable amount', 'taxable val'],
        tax: ['tax', 'tax amount', 'gst', 'tax value', 'gst amount'],
        qty: ['qty', 'quantity', 'qnty', 'units', 'net qty', 'net prod qty', 'total qty', 'sale qty', 'sales qty', 'pcs'],
        discount: ['direct discount', 'discount', 'total discount', 'disc', 'disc%'],
        indDiscount: ['indirect discount', 'ind discount'],
        dbdCharge: ['dbd charge', 'dbd'],
        procCharge: ['processing charge', 'proc charge'],
        svcCharge: ['service charge', 'svc charge'],
        addition: ['addition', 'additions'],
        deduction: ['deduction', 'deductions'],
        invoice: ['invoice number', 'invoice no', 'invoice', 'bill no', 'bill number', 'bill no.', 'invoice no.', 'inv no', 'receipt no'],
        customerName: ['customer name', 'customer', 'cust name', 'buyer name', 'buyer', 'party name', 'party', 'client name', 'client'],
        customerNo: ['customer number', 'customer no', 'cust no', 'mobile', 'phone', 'contact', 'mobile no', 'phone no', 'contact no', 'customer mobile', 'cust mobile', 'mobile number'],
    };

    // ---- COLUMN MAPPING (OSG file) ----
    const OSG_COL_MAP = {
        branch: ['store name', 'store', 'branch', 'branch name', 'outlet', 'outlet name', 'shop name'],
        storeCode: ['store code', 'store id', 'outlet code'],
        product: ['product', 'product name', 'product type', 'model', 'model name', 'item name', 'product group'],
        category: ['category', 'product category', 'item category'],
        brand: ['brand', 'brand name', 'make'],
        soldPrice: ['sold price', 'soldprice', 'plan price', 'selling price', 'net amount', 'amount', 'value', 'net value', 'sale price', 'mop', 'total amount', 'premium', 'premium amount'],
        qty: ['quantity', 'qty', 'ews qty', 'qnty', 'units', 'net qty', 'count', 'nos', 'pcs'],
        invoice: ['invoice no', 'invoice number', 'invoice', 'bill no', 'bill number', 'bill no.', 'invoice no.', 'inv no', 'receipt no'],
    };

    // ---- DOM REFERENCES ----
    const $ = id => document.getElementById(id);
    const sidebar = $('sidebar');
    const menuToggle = $('menuToggle');
    const loadingOverlay = $('loadingOverlay');
    const fileCountBadge = $('fileCountBadge');
    const fileCountText = $('fileCountText');
    const btnShare = $('btnShare');
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

    // ---- AUTHENTICATION STATE ----
    let isAuthenticated = false;
    const AUTH_PASSWORD = 'user1234';
    let pendingNavTarget = null;
    let pendingNavItem = null;

    // ---- CUSTOMER CALL TRACKING STATE ----
    const coStatusMap = {};
    const CO_CALLERS = [
        { name: 'Harmiya',  color: '#7c3aed', bg: 'rgba(124,58,237,0.15)' },
        { name: 'Aswathi',  color: '#0891b2', bg: 'rgba(8,145,178,0.15)' },
        { name: 'Shikha',   color: '#d97706', bg: 'rgba(217,119,6,0.15)'  },
    ];
    let currentCaller = localStorage.getItem('co_caller') || null;
    let coCurrentRows = [];   // current filtered rows
    let coDisplayLimit = 100; // pagination: rows shown so far

    // ---- GLOBAL DOWNLOAD STAFF DETAILS ----
    window.downloadStaffDetails = function(staff, branch) {
        if (!productData || productData.length === 0) return;
        
        let pRows = productData;
        if (staff) pRows = pRows.filter(r => r.staff === staff);
        if (branch) pRows = pRows.filter(r => r.branch === branch);
        
        if (pRows.length === 0) {
            alert('No detailed product data found.');
            return;
        }

        const osgInvoices = new Set();
        (osgData || []).forEach(r => { if (r.invoice) osgInvoices.add(r.invoice); });

        // Calculate staff-level overall conversions and product counts
        const staffInvoices = {};
        const staffStats = {};
        const staffProdCounts = {};

        pRows.forEach(r => {
            const s = r.staff || 'Unknown';
            const p = r.product || 'Unknown';

            if (!staffInvoices[s]) staffInvoices[s] = new Set();
            if (r.invoice) staffInvoices[s].add(r.invoice);
            
            if (!staffStats[s]) staffStats[s] = { pQty: 0, pRev: 0, oQty: 0, oRev: 0 };
            staffStats[s].pQty += r.qty || 0;
            staffStats[s].pRev += r.soldPrice || 0;

            if (!staffProdCounts[s]) staffProdCounts[s] = {};
            if (!staffProdCounts[s][p]) staffProdCounts[s][p] = 0;
            staffProdCounts[s][p] += (r.qty || 0);
        });

        (osgData || []).forEach(r => {
            if (!r.invoice) return;
            for (const s in staffInvoices) {
                if (staffInvoices[s].has(r.invoice)) {
                    staffStats[s].oQty += r.qty || 0;
                    staffStats[s].oRev += r.soldPrice || 0;
                    break;
                }
            }
        });

        const lines = [];

        const hdr = ['Staff', 'Branch', 'Staff Qty Conv%', 'Staff Val Conv%', 'Product', 'Total Sold Qty'];
        lines.push(hdr.join(','));
        
        Object.keys(staffProdCounts).sort().forEach(s => {
            const st = staffStats[s];
            const qConv = st.pQty > 0 ? (st.oQty / st.pQty) * 100 : 0;
            const vConv = st.pRev > 0 ? (st.oRev / st.pRev) * 100 : 0;
            
            // Find the branch for this staff from the filtered rows
            const sBranch = pRows.find(r => r.staff === s)?.branch || branch || 'Unknown';

            const products = Object.keys(staffProdCounts[s]).sort((a, b) => staffProdCounts[s][b] - staffProdCounts[s][a]);
            products.forEach(p => {
                lines.push([
                    q(s), q(sBranch), 
                    qConv.toFixed(2) + '%', vConv.toFixed(2) + '%', 
                    q(p), staffProdCounts[s][p]
                ].join(','));
            });
        });

        const filename = staff ? `details_${staff.replace(/[^a-z0-9]/gi, '_')}.csv` : `details_${branch.replace(/[^a-z0-9]/gi, '_')}.csv`;
        downloadCSV(lines.join('\n'), filename);
    };

    window.downloadBranchDetails = function(branch) {
        window.downloadStaffDetails(null, branch);
    };

    // ---- NAVIGATION ----
    document.querySelectorAll('.nav-item').forEach(item => {
        item.addEventListener('click', (e) => {
            e.preventDefault();
            const section = item.dataset.section;

            // Check if page needs auth — Upload Data and Customers Without OSG are public
            const isPublicPage = section === 'customers-osg-section' || section === 'upload-section';
            if (!isPublicPage && !isAuthenticated) {
                // Intercept navigation and show password modal
                pendingNavTarget = section;
                pendingNavItem = item;
                $('passwordModal').style.display = 'flex';
                $('modalPasswordInput').value = '';
                $('passwordErrorMsg').textContent = '';
                $('modalPasswordInput').focus();
                return;
            }

            performNavigation(item, section);
        });
    });

    function performNavigation(item, section) {
        document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
        item.classList.add('active');
        document.querySelectorAll('.content-section').forEach(s => s.classList.remove('active'));
        $(section).classList.add('active');
        $('pageTitle').textContent = item.querySelector('span').textContent;
        sidebar.classList.remove('open');
    }

    // ---- PASSWORD MODAL LOGIC ----
    $('btnSubmitPassword').addEventListener('click', handlePasswordSubmit);
    $('btnCancelPassword').addEventListener('click', () => {
        $('passwordModal').style.display = 'none';
        pendingNavTarget = null;
        pendingNavItem = null;
    });
    $('modalPasswordInput').addEventListener('keypress', (e) => {
        if (e.key === 'Enter') handlePasswordSubmit();
    });

    function handlePasswordSubmit() {
        const input = $('modalPasswordInput').value;
        if (input === AUTH_PASSWORD) {
            isAuthenticated = true;
            $('passwordModal').style.display = 'none';
            if (pendingNavItem && pendingNavTarget) {
                performNavigation(pendingNavItem, pendingNavTarget);
            }
        } else {
            $('passwordErrorMsg').textContent = 'Incorrect password. Try again.';
            $('modalPasswordInput').classList.add('shake');
            setTimeout(() => $('modalPasswordInput').classList.remove('shake'), 400);
        }
    }
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
        zone.addEventListener('click', (e) => {
            if (e.target === input) return;
            input.click();
        });
        input.addEventListener('click', (e) => e.stopPropagation());
        zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag-over'); });
        zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
        zone.addEventListener('drop', async e => {
            e.preventDefault();
            zone.classList.remove('drag-over');
            if (e.dataTransfer.files.length > 0) {
                showLoading(true);
                try {
                    await onFile(e.dataTransfer.files[0]);
                } catch (err) {
                    console.error(err);
                } finally {
                    showLoading(false);
                    // Reset input value in drop to allow re-upload of same file via click later
                    input.value = '';
                }
            }
        });
        input.addEventListener('change', async () => {
            if (input.files.length > 0) {
                showLoading(true);
                try {
                    await onFile(input.files[0]);
                } catch (err) {
                    console.error(err);
                } finally {
                    showLoading(false);
                    // Reset input value to allow re-upload of the same file
                    input.value = '';
                }
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

    btnGenerate.addEventListener('click', () => {
        showLoading(true);
        
        // Use setTimeout to yield the main thread allowing the loading UI to render before heavy processing
        setTimeout(() => {
            try {
                allData = [...productData, ...amcData];

                fileCountBadge.style.display = 'flex';
                fileCountText.textContent = `${allData.length} product · ${osgData.length} OSG`;
                btnShare.style.display = 'flex';
                btnReset.style.display = 'flex';

                populateFilters();
                applyFilters();

                document.querySelector('[data-section="dashboard-section"]').click();
            } catch (err) {
                console.error('[Generate Error]', err);
                alert('An error occurred while generating reports:\n' + err.message);
            } finally {
                showLoading(false);
            }
        }, 50);
    });

    // ---- SHARE DASHBOARD LOGIC ----
    btnShare.addEventListener('click', () => {
        if (productData.length === 0) return alert('Upload data first via Dashboard.');

        // Find missedUnique for the whole dataset
        const osgInvoices = new Set();
        osgData.forEach(r => { if (r.invoice) osgInvoices.add(r.invoice); });
        const seenInv = new Set();
        const fullMissedUnique = [];
        productData.forEach(r => {
            if (r.invoice && !osgInvoices.has(r.invoice) && !seenInv.has(r.invoice)) {
                seenInv.add(r.invoice);
                fullMissedUnique.push(r);
            }
        });

        // Strip to only display fields, sort by value high-to-low, cap at 2000 top-priority customers
        const payload = fullMissedUnique
            .sort((a, b) => (b.soldPrice || 0) - (a.soldPrice || 0))
            .slice(0, 2000)
            .map(r => ({
                invoice:      r.invoice      || '',
                customerName: r.customerName || '',
                customerNo:   r.customerNo   || '',
                staff:        r.staff        || '',
                branch:       r.branch       || '',
                product:      r.product      || '',
                soldPrice:    r.soldPrice    || 0,
                qty:          r.qty          || 0,
            }));

        showLoading(true);
        try {
            const shareRef = firebase.database().ref('shares').push();
            shareRef.set({ missedUnique: payload, timestamp: Date.now() })
                .then(() => {
                    showLoading(false);
                    const base = window.location.protocol === 'file:' ? 'http://myg-analytics-2026.surge.sh/' : window.location.origin + window.location.pathname;
                    const shareUrl = base + '?share=' + shareRef.key;
                    if (navigator.clipboard) {
                        navigator.clipboard.writeText(shareUrl).then(() => {
                            alert('Share link copied to clipboard!\n\nLink: ' + shareUrl);
                        }).catch(() => alert('Share Link generated: \n\n' + shareUrl));
                    } else {
                        alert('Share Link generated: \n\n' + shareUrl);
                    }
                }).catch(e => {
                    showLoading(false);
                    alert('Failed to generate share link: ' + e.message);
                });
        } catch (err) {
            showLoading(false);
            console.error(err);
            alert('Firebase configuration error (likely missing databaseURL). Cannot share dashboard right now: ' + err.message);
        }
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
            r.customerName = strVal(row, mapping.customerName);
            r.customerNo = strVal(row, mapping.customerNo);
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
        // Detect CSV files and use fast CSV parser
        const isCSV = file.name.toLowerCase().endsWith('.csv');
        if (isCSV) return parseCSVFile(file, colMap, rowMapper);

        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const fileSizeMB = (data.length / (1024 * 1024)).toFixed(1);
                    updateLoadingMsg('Reading XLSX (' + fileSizeMB + ' MB)...');
                    console.log('[Excel] Loading ' + file.name + ' (' + fileSizeMB + ' MB)');

                    if (data.length > 20 * 1024 * 1024) {
                        alert('This XLSX file is ' + fileSizeMB + ' MB which is very large.\n\nFor best results, convert it to CSV first:\n1. Open a terminal in the portal folder\n2. Run: python convert.py "your_file.xlsx"\n3. Upload the resulting .csv file instead.\n\nWill attempt to parse anyway...');
                    }

                    setTimeout(() => {
                        try {
                            const wb = XLSX.read(data, { type: 'array', cellDates: true, dense: true });
                            let allRows = [];
                            for (let si = 0; si < wb.SheetNames.length; si++) {
                                const sheetName = wb.SheetNames[si];
                                updateLoadingMsg('Parsing sheet ' + (si + 1) + '/' + wb.SheetNames.length + '...');
                                const sheet = wb.Sheets[sheetName];
                                const json = XLSX.utils.sheet_to_json(sheet, { defval: '' });
                                if (json.length === 0) continue;
                                const headers = Object.keys(json[0]);
                                console.log('[Sheet: ' + sheetName + '] ' + json.length + ' rows');
                                const mapping = autoMapColumns(headers, colMap);
                                for (let i = 0; i < json.length; i++) {
                                    allRows.push(rowMapper(json[i], mapping));
                                }
                            }
                            if (allRows.length === 0) { reject(new Error('No data')); return; }
                            updateLoadingMsg('Loaded ' + allRows.length.toLocaleString() + ' rows!');
                            console.log('[Total] ' + allRows.length + ' rows');
                            resolve(allRows);
                        } catch (err) {
                            console.error('[Parse Error]', err);
                            alert('Error parsing XLSX: ' + err.message + '\n\nFor large files, convert to CSV first:\npython convert.py "your_file.xlsx"');
                            reject(err);
                        }
                    }, 100);
                } catch (err) { reject(err); }
            };
            reader.onerror = () => reject(new Error('File read error'));
            reader.readAsArrayBuffer(file);
        });
    }

    function parseCSVFile(file, colMap, rowMapper) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const text = e.target.result;
                    const fileSizeMB = (text.length / (1024 * 1024)).toFixed(1);
                    updateLoadingMsg('Parsing CSV (' + fileSizeMB + ' MB)...');
                    console.log('[CSV] Loading ' + file.name + ' (' + fileSizeMB + ' MB)');

                    // Use XLSX library to parse CSV (handles all edge cases)
                    var wb = XLSX.read(text, { type: 'string' });
                    var sheet = wb.Sheets[wb.SheetNames[0]];
                    var json = XLSX.utils.sheet_to_json(sheet, { defval: '' });

                    console.log('[CSV] ' + json.length + ' rows parsed');
                    if (json.length === 0) { reject(new Error('No data in CSV')); return; }

                    var headers = Object.keys(json[0]);
                    console.log('[CSV Headers] ' + headers.slice(0, 8).join(', '));
                    var mapping = autoMapColumns(headers, colMap);

                    updateLoadingMsg('Processing ' + json.length.toLocaleString() + ' rows...');
                    var allRows = [];
                    for (var i = 0; i < json.length; i++) {
                        allRows.push(rowMapper(json[i], mapping));
                        if (i % 100000 === 0 && i > 0) {
                            updateLoadingMsg('Row ' + i.toLocaleString() + ' / ' + json.length.toLocaleString());
                        }
                    }

                    updateLoadingMsg('Loaded ' + allRows.length.toLocaleString() + ' rows!');
                    console.log('[CSV Total] ' + allRows.length + ' rows');
                    resolve(allRows);
                } catch (err) {
                    console.error('[CSV Error]', err);
                    alert('Error parsing CSV: ' + err.message);
                    reject(err);
                }
            };
            reader.onerror = () => reject(new Error('File read error'));
            reader.readAsText(file);
        });
    }

    function updateLoadingMsg(msg) {
        const el = document.querySelector('#loadingOverlay p');
        if (el) el.textContent = msg;
    }

    function autoMapColumns(headers, colMap) {
        const mapping = {};
        const headersLower = headers.map(h => h.toLowerCase().trim());

        for (const [key, aliases] of Object.entries(colMap)) {
            mapping[key] = null;

            // Pass 1: Exact match (alias order = priority)
            for (const alias of aliases) {
                const idx = headersLower.indexOf(alias);
                if (idx >= 0) { mapping[key] = headers[idx]; break; }
            }

            // Pass 2: Partial/contains match as fallback
            if (!mapping[key]) {
                for (const alias of aliases) {
                    const idx = headersLower.findIndex(h => h.includes(alias) || alias.includes(h));
                    if (idx >= 0 && h_len(headersLower[idx]) > 1) { mapping[key] = headers[idx]; break; }
                }
            }
        }

        // Warn about critical unmapped columns
        const critical = ['soldPrice', 'qty', 'branch', 'product'];
        critical.forEach(k => {
            if (!mapping[k]) console.warn(`[⚠ Column NOT FOUND] '${k}' — no matching header. Available headers:`, headers.join(', '));
        });

        console.log('[Column Mapping]', JSON.stringify(mapping));
        return mapping;
    }

    function h_len(s) { return s ? s.length : 0; }

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
        sel.addEventListener('change', () => {
            try {
                applyFilters();
            } catch (err) {
                console.error('[Filter Error]', err);
                alert('An error occurred while applying filters:\n' + err.message);
            }
        });
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
        green: '#10b981',   // Emerald 500
        red: '#ef4444',     // Rose 500
        blue: '#3b82f6',    // Cobalt 500
        amber: '#f59e0b',   // Amber 500
        cyan: '#06b6d4',    // Cyan 500
        purple: '#8b5cf6',  // Violet 500
        orange: '#f97316',  // Orange 500
        slate: '#64748b'    // Slate 500
    };
    const PALETTE = [
        '#3b82f6', '#10b981', '#8b5cf6', '#f59e0b', '#06b6d4', '#f43f5e', 
        '#8b5cf6', '#14b8a6', '#f97316', '#6366f1', '#0ea5e9'
    ];

    function chartBase(withYCurrency) {
        return {
            responsive: true, 
            maintainAspectRatio: false,
            plugins: { 
                legend: { 
                    labels: { 
                        color: '#94a3b8', 
                        font: { family: "'Outfit', sans-serif", size: 12, weight: '500' },
                        padding: 20,
                        usePointStyle: true
                    } 
                },
                tooltip: {
                    backgroundColor: '#1e293b',
                    titleColor: '#f1f5f9',
                    bodyColor: '#cbd5e1',
                    borderColor: 'rgba(255,255,255,0.1)',
                    borderWidth: 1,
                    padding: 12,
                    cornerRadius: 8,
                    displayColors: true
                }
            },
            scales: {
                x: { 
                    ticks: { color: '#64748b', font: { family: "'Outfit', sans-serif", size: 11 } }, 
                    grid: { color: 'rgba(255,255,255,0.03)', drawBorder: false } 
                },
                y: { 
                    ticks: { 
                        color: '#64748b', 
                        font: { family: "'Outfit', sans-serif", size: 11 }, 
                        callback: withYCurrency ? v => fmtShort(v) : v => v + '%' 
                    }, 
                    grid: { color: 'rgba(255,255,255,0.03)', drawBorder: false } 
                },
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
        $('fsTableWrapper').innerHTML = noDataHTML('Upload data and generate reports first.');
        $('fsKpiRow').innerHTML = '';
        $('fsCount').textContent = '0 staff';
        $('pdTopRevTable').innerHTML = '';
        $('pdTopConvTable').innerHTML = '';
        $('pdKpiRow').innerHTML = '';
        $('coMissedTable').innerHTML = noDataHTML('Upload data and generate reports first.');
        $('coMissedCount').textContent = '0 customers';
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
        const selRBM = $('lcRBM').value;
        const selBDM = $('lcBDM').value;

        const allStats = buildStaffStats();

        // Populate RBM dropdown (preserve selection)
        const rbmSet = [...new Set(allStats.map(s => s.rbm).filter(Boolean))].sort();
        const bdmSet = [...new Set(allStats.map(s => s.bdm).filter(Boolean))].sort();

        const rbmEl = $('lcRBM');
        const bdmEl = $('lcBDM');
        const prevRBM = selRBM;
        const prevBDM = selBDM;

        rbmEl.innerHTML = '<option value="">All RBMs</option>' +
            rbmSet.map(r => `<option value="${r}" ${r === prevRBM ? 'selected' : ''}>${r}</option>`).join('');
        bdmEl.innerHTML = '<option value="">All BDMs</option>' +
            bdmSet.map(b => `<option value="${b}" ${b === prevBDM ? 'selected' : ''}>${b}</option>`).join('');

        // Filter: minQty, maxConv, optional RBM, optional BDM
        const filtered = allStats
            .filter(s => s.pQty >= minQty && s.qtyConv <= maxConv)
            .filter(s => !selRBM || s.rbm === selRBM)
            .filter(s => !selBDM || s.bdm === selBDM)
            .sort((a, b) => {
                if (a.qtyConv !== b.qtyConv) return a.qtyConv - b.qtyConv;
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
            const dlIcon = `<button onclick="window.downloadStaffDetails('${e.name}', '${e.branch}')" title="Download Staff Details" style="background:none;border:none;cursor:pointer;color:var(--primary);padding:0;margin-left:8px;vertical-align:middle;display:inline-flex;align-items:center;"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg></button>`;
            html += `<tr>
                <td class="number-cell">${rankBadge}</td>
                <td style="white-space:nowrap;"><strong>${e.name}</strong>${dlIcon}</td>
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
        const selRBM = $('lcRBM').value;
        const selBDM = $('lcBDM').value;
        const allStats = buildStaffStats();
        const filtered = allStats
            .filter(s => s.pQty >= minQty && s.qtyConv <= maxConv)
            .filter(s => !selRBM || s.rbm === selRBM)
            .filter(s => !selBDM || s.bdm === selBDM)
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
        const selRBM = $('tcRBM').value;
        const selBDM = $('tcBDM').value;

        const allStats = buildStaffStats();

        // Populate RBM and BDM dropdowns (preserve selection)
        const rbmSet = [...new Set(allStats.map(s => s.rbm).filter(Boolean))].sort();
        const bdmSet = [...new Set(allStats.map(s => s.bdm).filter(Boolean))].sort();
        $('tcRBM').innerHTML = '<option value="">All RBMs</option>' +
            rbmSet.map(r => `<option value="${r}" ${r === selRBM ? 'selected' : ''}>${r}</option>`).join('');
        $('tcBDM').innerHTML = '<option value="">All BDMs</option>' +
            bdmSet.map(b => `<option value="${b}" ${b === selBDM ? 'selected' : ''}>${b}</option>`).join('');

        // Filter: must have >= minQty product qty AND conversion > 0, plus RBM/BDM filters
        const eligible = allStats
            .filter(s => s.pQty >= minQty && s[sortBy] > 0)
            .filter(s => !selRBM || s.rbm === selRBM)
            .filter(s => !selBDM || s.bdm === selBDM);

        // Sort: primarily by absolute OSG volume (OSG Qty or OSG Revenue)
        // This ensures staff with the highest actual number of conversions are at the top,
        // which naturally requires both high Product Qty and high Conversion %.
        const filtered = eligible
            .sort((a, b) => {
                const volA = sortBy === 'qtyConv' ? a.oQty : a.oRev;
                const volB = sortBy === 'qtyConv' ? b.oQty : b.oRev;
                if (volB !== volA) return volB - volA; // Highest OSG volume first
                return b[sortBy] - a[sortBy];          // Tie-breaker: highest conversion %
            })
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
            const dlIcon = `<button onclick="window.downloadStaffDetails('${e.name}', '${e.branch}')" title="Download Staff Details" style="background:none;border:none;cursor:pointer;color:var(--primary);padding:0;margin-left:8px;vertical-align:middle;display:inline-flex;align-items:center;"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg></button>`;
            html += `<tr>
                <td class="number-cell">${rankBadge}</td>
                <td style="white-space:nowrap;"><strong>${e.name}</strong>${dlIcon}</td>
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
        const selRBM = $('tcRBM').value;
        const selBDM = $('tcBDM').value;
        const allStats = buildStaffStats();
        const filtered = allStats
            .filter(s => s.pQty >= minQty && s[sortBy] > 0)
            .filter(s => !selRBM || s.rbm === selRBM)
            .filter(s => !selBDM || s.bdm === selBDM)
            .sort((a, b) => {
                const volA = sortBy === 'qtyConv' ? a.oQty : a.oRev;
                const volB = sortBy === 'qtyConv' ? b.oQty : b.oRev;
                if (volB !== volA) return volB - volA; // Highest OSG volume first
                return b[sortBy] - a[sortBy];          // Tie-breaker: highest conversion %
            })
            .slice(0, topN);
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

        // ---- Card X: Deep Root Cause Analysis ----
        let deepAnalysisHtml = `<ul style="padding-left:1.5rem; margin-bottom:1rem; color:var(--text-secondary);">`;

        // 1. Product factor
        const lowConvProds = prodStats
            .filter(p => p.pQty >= 5) // At least 5 units sold
            .sort((a, b) => a.qtyConv - b.qtyConv);

        if (lowConvProds.length > 0) {
            const worstProd = lowConvProds[0];
            deepAnalysisHtml += `
                <li style="margin-bottom:12px;">
                    <strong>Product Factor (<span style="color:var(--loss);">${worstProd.name}</span>):</strong> 
                    Only converting at <span style="color:var(--loss); font-weight:600;">${worstProd.qtyConv.toFixed(1)}%</span>.
                    <div style="margin-top:4px; font-size:0.9rem;">
                        <em>Reason:</em> Highly competitive market segment or low perceived value of OSG for this specific product tier. Customers might view the base product as disposable or already adequately warrantied by the manufacturer.
                    </div>
                    <div style="margin-top:4px; font-size:0.9rem;">
                        <em>Recommendation:</em> Bundle OSG directly into the financing plan for this product, or introduce a "lite" OSG tier tailored to its price point.
                    </div>
                </li>`;
        }

        // 2. Staff factor
        const weakStaff = staffStats.filter(s => s.pQty >= 5 && s.qtyConv > 0 && s.qtyConv < 5);
        if (zeroConvStaff.length > 0 || weakStaff.length > 0) {
            const problemStaffCount = zeroConvStaff.length + weakStaff.length;
            const pctStaff = ((problemStaffCount / staffStats.length) * 100).toFixed(0);
            deepAnalysisHtml += `
                <li style="margin-bottom:12px;">
                    <strong>Staff Effectiveness:</strong> 
                    <span style="color:var(--loss); font-weight:600;">${pctStaff}%</span> of staff (${problemStaffCount} members) are significantly underperforming (&lt;5% conversion).
                    <div style="margin-top:4px; font-size:0.9rem;">
                        <em>Reason:</em> Lack of pitch confidence, skipping the OSG conversation entirely to close the primary sale faster, or failing to overcome initial customer objections ("it's too expensive").
                    </div>
                    <div style="margin-top:4px; font-size:0.9rem;">
                        <em>Recommendation:</em> Implement a mandatory "3-strike objection handling" framework. Pair bottom quartile staff with Top 3 converters (${topQty.slice(0, 3).map(s => s.name).join(', ')}) for mandatory shadowing sessions this week.
                    </div>
                </li>`;
        }

        // 3. Branch/Footprint factor
        if (weakBranches.length > 0) {
            const worstBranch = weakBranches.sort((a, b) => a.qtyConv - b.qtyConv)[0];
            deepAnalysisHtml += `
                <li style="margin-bottom:12px;">
                    <strong>Location Impact (<span style="color:var(--loss);">${worstBranch.name}</span>):</strong> 
                    Converting at just <span style="color:var(--loss); font-weight:600;">${worstBranch.qtyConv.toFixed(1)}%</span>.
                    <div style="margin-top:4px; font-size:0.9rem;">
                        <em>Reason:</em> Demographic price sensitivity in this catchment area, or systemic store-leadership de-prioritization of accessory/OSG targets in favor of raw hardware volume.
                    </div>
                    <div style="margin-top:4px; font-size:0.9rem;">
                        <em>Recommendation:</em> Run a weekend "Free Protection Demo" in-store. RBM needs to reset expectations with the Store Manager that hardware quotas include an absolute minimum 15% OSG attach.
                    </div>
                </li>`;
        }
        deepAnalysisHtml += `</ul>`;

        html += insightCard('🔍', 'Deep Root Cause Analysis', 'danger', `
            <p style="margin-bottom:1rem; color:var(--text-primary); font-weight:500;">Based on combinatorial data analysis, the primary drivers of lost conversion are:</p>
            ${deepAnalysisHtml}
        `);

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

    // ---- FUTURE STORES PAGE ----
    $('btnFSRefresh').addEventListener('click', renderFutureStoresPage);
    $('btnFSExport').addEventListener('click', exportFutureStoresCSV);
    document.querySelector('[data-section="future-section"]').addEventListener('click', () => {
        setTimeout(renderFutureStoresPage, 50);
    });

    function buildFutureStaffStats() {
        // Only include product rows whose branch contains "FUTURE" (case-insensitive)
        const futureProduct = productData.filter(r => r.branch && r.branch.toUpperCase().includes('FUTURE'));

        // Build invoice → staff lookup from future product data
        const invoiceStaff = {};
        futureProduct.forEach(r => { if (r.invoice && r.staff) invoiceStaff[r.invoice] = r.staff; });

        // Group product data by staff
        const pByStaff = {};
        futureProduct.forEach(r => {
            const s = r.staff || 'Unknown';
            if (!pByStaff[s]) pByStaff[s] = { branch: r.branch, rbm: r.rbm, bdm: r.bdm, rows: [] };
            pByStaff[s].rows.push(r);
        });

        // Group OSG data by staff via invoice mapping (only future-related invoices)
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
            
            const prodCounts = {};
            pInfo.rows.forEach(r => {
                const p = r.product || 'Unknown';
                prodCounts[p] = (prodCounts[p] || 0) + (r.qty || 0);
            });
            const oProdCounts = {};
            oRows.forEach(r => {
                const p = r.product || 'Unknown';
                oProdCounts[p] = (oProdCounts[p] || 0) + (r.qty || 0);
            });
            const allProds = new Set([...Object.keys(prodCounts), ...Object.keys(oProdCounts)]);
            const products = Array.from(allProds).map(p => ({
                name: p,
                qty: prodCounts[p] || 0,
                osgQty: oProdCounts[p] || 0
            })).sort((a,b) => b.qty - a.qty);

            return { name, branch: pInfo.branch, rbm: pInfo.rbm, bdm: pInfo.bdm, pQty, oQty, pRev, oRev, qtyConv, valConv, products };
        });
    }

    function renderFutureStoresPage() {
        if (productData.length === 0) {
            $('fsTableWrapper').innerHTML = noDataHTML('Upload data and generate reports first.');
            $('fsKpiRow').innerHTML = '';
            $('fsCount').textContent = '0 staff';
            return;
        }

        const selRBM = $('fsRBM').value;
        const selBDM = $('fsBDM').value;
        const selBranch = $('fsBranch').value;

        const allStats = buildFutureStaffStats();

        // Populate dropdowns (preserve selection)
        const rbmSet = [...new Set(allStats.map(s => s.rbm).filter(Boolean))].sort();
        const bdmSet = [...new Set(allStats.map(s => s.bdm).filter(Boolean))].sort();
        const branchSet = [...new Set(allStats.map(s => s.branch).filter(Boolean))].sort();

        $('fsRBM').innerHTML = '<option value="">All RBMs</option>' +
            rbmSet.map(r => `<option value="${r}" ${r === selRBM ? 'selected' : ''}>${r}</option>`).join('');
        $('fsBDM').innerHTML = '<option value="">All BDMs</option>' +
            bdmSet.map(b => `<option value="${b}" ${b === selBDM ? 'selected' : ''}>${b}</option>`).join('');
        $('fsBranch').innerHTML = '<option value="">All Future Stores</option>' +
            branchSet.map(b => `<option value="${b}" ${b === selBranch ? 'selected' : ''}>${b}</option>`).join('');

        // Filter
        const filtered = allStats
            .filter(s => !selRBM || s.rbm === selRBM)
            .filter(s => !selBDM || s.bdm === selBDM)
            .filter(s => !selBranch || s.branch === selBranch)
            .sort((a, b) => b.pQty - a.pQty);

        $('fsCount').textContent = `${filtered.length} staff`;

        // KPI summary
        const totalPQty = filtered.reduce((s, r) => s + r.pQty, 0);
        const totalOQty = filtered.reduce((s, r) => s + r.oQty, 0);
        const totalPRev = filtered.reduce((s, r) => s + r.pRev, 0);
        const totalORev = filtered.reduce((s, r) => s + r.oRev, 0);
        const avgQtyConv = totalPQty > 0 ? (totalOQty / totalPQty) * 100 : 0;
        const avgValConv = totalPRev > 0 ? (totalORev / totalPRev) * 100 : 0;
        const uniqueBranches = new Set(filtered.map(s => s.branch)).size;

        $('fsKpiRow').innerHTML = `
            <div class="lc-kpi"><span class="lc-kpi-label">Future Stores</span><span class="lc-kpi-val">${uniqueBranches}</span></div>
            <div class="lc-kpi"><span class="lc-kpi-label">Total Staff</span><span class="lc-kpi-val">${filtered.length}</span></div>
            <div class="lc-kpi"><span class="lc-kpi-label">Avg Qty Conv</span><span class="lc-kpi-val conversion-text">${avgQtyConv.toFixed(2)}%</span></div>
            <div class="lc-kpi"><span class="lc-kpi-label">Avg Val Conv</span><span class="lc-kpi-val conversion-text">${avgValConv.toFixed(2)}%</span></div>
            <div class="lc-kpi"><span class="lc-kpi-label">Total Product Revenue</span><span class="lc-kpi-val">${fmtShort(totalPRev)}</span></div>
            <div class="lc-kpi"><span class="lc-kpi-label">Total OSG Revenue</span><span class="lc-kpi-val">${fmtShort(totalORev)}</span></div>
        `;

        if (filtered.length === 0) {
            $('fsTableWrapper').innerHTML = noDataHTML('No staff found in Future stores.');
            return;
        }

        let html = `<table class="data-table">
            <thead><tr>
                <th>#</th><th>Staff</th><th>Branch</th><th>RBM</th><th>BDM</th>
                <th>Prod Qty</th><th>OSG Qty</th><th>Qty Conv%</th><th>Val Conv%</th><th>Prod Rev</th><th>OSG Rev</th>
            </tr></thead><tbody>`;

        filtered.forEach((e, i) => {
            const convCls = e.qtyConv === 0 ? 'loss-val' : (e.qtyConv < 5 ? 'conv-warn' : 'conv-val');
            const rank = i + 1;
            const rankBadge = rank <= 3 ? `<span class="rank-badge rank-${rank}">${rank}</span>` : `<span class="rank-num">${rank}</span>`;
            const dlIcon = `<button onclick="window.downloadStaffDetails('${e.name}', '${e.branch}')" title="Download Staff Details" style="background:none;border:none;cursor:pointer;color:var(--primary);padding:0;margin-left:8px;vertical-align:middle;display:inline-flex;align-items:center;"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg></button>`;
            const dlBranchIcon = `<button onclick="window.downloadBranchDetails('${e.branch}')" title="Download Branch Details" style="background:none;border:none;cursor:pointer;color:var(--primary);padding:0;margin-left:8px;vertical-align:middle;display:inline-flex;align-items:center;"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg></button>`;
            html += `<tr>
                <td class="number-cell">${rankBadge}</td>
                <td style="white-space:nowrap;"><strong>${e.name}</strong>${dlIcon}</td>
                <td style="white-space:nowrap;">${e.branch}${dlBranchIcon}</td>
                <td>${e.rbm}</td>
                <td>${e.bdm}</td>
                <td class="number-cell"><strong>${e.pQty}</strong></td>
                <td class="number-cell">${e.oQty}</td>
                <td class="number-cell ${convCls}">${e.qtyConv.toFixed(2)}%</td>
                <td class="number-cell conv-val">${e.valConv.toFixed(2)}%</td>
                <td class="number-cell">${fmtShort(e.pRev)}</td>
                <td class="number-cell">${fmtShort(e.oRev)}</td>
            </tr>`;
        });

        html += '</tbody></table>';
        $('fsTableWrapper').innerHTML = html;
    }

    function exportFutureStoresCSV() {
        if (productData.length === 0) return;
        const selRBM = $('fsRBM').value;
        const selBDM = $('fsBDM').value;
        const selBranch = $('fsBranch').value;
        const allStats = buildFutureStaffStats();
        const filtered = allStats
            .filter(s => !selRBM || s.rbm === selRBM)
            .filter(s => !selBDM || s.bdm === selBDM)
            .filter(s => !selBranch || s.branch === selBranch)
            .sort((a, b) => b.pQty - a.pQty);
        if (filtered.length === 0) return;
        if (filtered.length === 0) return;
        const hdrAll    = ['Rank', 'Staff', 'Product', 'Product Qty', 'Branch', 'RBM', 'BDM', 'Total Prod Qty', 'OSG Qty', 'Qty Conv%', 'Val Conv%'];
        const hdrBranch = ['Rank', 'Staff', 'Product', 'Total Prod Qty', 'Prod Qty', 'OSG Qty', 'Qty Conv%', 'Val Conv%'];
        const wb = XLSX.utils.book_new();

        function addSheet(statsList, sheetName, isBranchSheet) {
            const hdr = isBranchSheet ? hdrBranch : hdrAll;
            const data = [hdr];
            const mergeRanges = [];
            // For All: Rank(0), Staff(1), Branch(4), RBM(5), BDM(6), TotQty(7), QtyConv(9), ValConv(10)
            // For Branch: Rank(0), Staff(1), TotQty(3), QtyConv(6), ValConv(7)
            const mergeCols = isBranchSheet ? [0, 1, 3, 6, 7] : [0, 1, 4, 5, 6, 7, 9, 10];

            statsList.forEach((e, i) => {
                const rank = i + 1;
                const startRow = data.length;

                if (e.products && e.products.length > 0) {
                    e.products.forEach((prod, pIdx) => {
                        if (pIdx === 0) {
                            if (isBranchSheet) {
                                // [Rank, Staff, Product, Total Prod Qty, Prod Qty, OSG Qty, Qty Conv%, Val Conv%]
                                data.push([rank, e.name, prod.name, e.pQty, prod.qty, prod.osgQty, parseFloat(e.qtyConv.toFixed(2)), parseFloat(e.valConv.toFixed(2))]);
                            } else {
                                // [Rank, Staff, Product, Product Qty, Branch, RBM, BDM, Total Prod Qty, OSG Qty, Qty Conv%, Val Conv%]
                                data.push([rank, e.name, prod.name, prod.qty, e.branch, e.rbm, e.bdm, e.pQty, prod.osgQty, parseFloat(e.qtyConv.toFixed(2)), parseFloat(e.valConv.toFixed(2))]);
                            }
                        } else {
                            if (isBranchSheet) {
                                data.push(['', '', prod.name, '', prod.qty, prod.osgQty, '', '']);
                            } else {
                                data.push(['', '', prod.name, prod.qty, '', '', '', '', prod.osgQty, '', '']);
                            }
                        }
                    });
                    const endRow = data.length - 1;
                    if (e.products.length > 1) {
                        mergeCols.forEach(c => mergeRanges.push({ s: { r: startRow, c }, e: { r: endRow, c } }));
                    }
                } else {
                    if (isBranchSheet) {
                        data.push([rank, e.name, '', e.pQty, '', e.oQty, parseFloat(e.qtyConv.toFixed(2)), parseFloat(e.valConv.toFixed(2))]);
                    } else {
                        data.push([rank, e.name, '', '', e.branch, e.rbm, e.bdm, e.pQty, e.oQty, parseFloat(e.qtyConv.toFixed(2)), parseFloat(e.valConv.toFixed(2))]);
                    }
                }
            });

            const ws = XLSX.utils.aoa_to_sheet(data);
            if (mergeRanges.length > 0) ws['!merges'] = mergeRanges;

            const headerStyle = {
                font: { bold: true, color: { rgb: 'FFFFFF' } },
                fill: { fgColor: { rgb: '3b82f6' } },
                alignment: { horizontal: 'center', vertical: 'center' }
            };
            const altRowStyle = { fill: { fgColor: { rgb: 'f1f5f9' } } };

            for (let R = 0; R < data.length; ++R) {
                for (let C = 0; C < hdr.length; ++C) {
                    const cell_ref = XLSX.utils.encode_cell({ c: C, r: R });
                    if (!ws[cell_ref]) ws[cell_ref] = { t: 's', v: '' };
                    if (R === 0) {
                        ws[cell_ref].s = headerStyle;
                    } else {
                        const baseStyle = R % 2 === 0 ? { ...altRowStyle } : {};
                        ws[cell_ref].s = mergeCols.includes(C)
                            ? { ...baseStyle, alignment: { horizontal: 'center', vertical: 'center' } }
                            : baseStyle;
                    }
                }
            }

            ws['!cols'] = isBranchSheet
                ? [{wch:8},{wch:25},{wch:22},{wch:14},{wch:10},{wch:10},{wch:12},{wch:12}]
                : [{wch:8},{wch:25},{wch:22},{wch:12},{wch:20},{wch:15},{wch:15},{wch:10},{wch:10},{wch:12},{wch:12}];

            let safeName = sheetName.substring(0, 31).replace(/[\\\/?*[\]]/g, '');
            if (!safeName) safeName = 'Sheet';
            let finalName = safeName;
            let counter = 1;
            while (wb.SheetNames.includes(finalName)) {
                finalName = safeName.substring(0, 27) + ' ' + counter;
                counter++;
            }
            XLSX.utils.book_append_sheet(wb, ws, finalName);
        }

        // First sheet: All (no change, detailed itemized)
        addSheet(filtered, 'future_stores_staff', false);

        // Branch sheets: Simplified format (8 cols)
        const branches = [...new Set(filtered.map(s => s.branch).filter(Boolean))].sort();
        branches.forEach(branch => {
            addSheet(filtered.filter(s => s.branch === branch), branch, true);
        });

        XLSX.writeFile(wb, 'future_stores_staff.xlsx');
    }

    // ---- PRODUCT DETAILS PAGE ----
    $('btnPDRefresh').addEventListener('click', renderProductDetailsPage);
    document.querySelector('[data-section="productdetails-section"]').addEventListener('click', () => {
        setTimeout(renderProductDetailsPage, 50);
    });

    // Auto-refresh when any Product Details dropdown changes
    ['pdRBM', 'pdBDM', 'pdProduct'].forEach(id => {
        $(id).addEventListener('change', renderProductDetailsPage);
    });

    function renderProductDetailsPage() {
        if (productData.length === 0) {
            $('pdTopRevTable').innerHTML = noDataHTML('Upload data and generate reports first.');
            $('pdTopConvTable').innerHTML = '';
            $('pdMissedTable').innerHTML = noDataHTML('Upload data and generate reports first.');
            $('pdKpiRow').innerHTML = '';
            $('pdMissedCount').textContent = '0 customers';
            return;
        }

        const selRBM = $('pdRBM').value;
        const selBDM = $('pdBDM').value;
        const selProduct = $('pdProduct').value;

        // Populate filter dropdowns
        const rbmSet = [...new Set(productData.map(r => r.rbm).filter(Boolean))].sort();
        const bdmSet = [...new Set(productData.map(r => r.bdm).filter(Boolean))].sort();
        const prodSet = [...new Set(productData.map(r => r.product).filter(Boolean))].sort();
        $('pdRBM').innerHTML = '<option value="">All RBMs</option>' +
            rbmSet.map(r => `<option value="${r}" ${r === selRBM ? 'selected' : ''}>${r}</option>`).join('');
        $('pdBDM').innerHTML = '<option value="">All BDMs</option>' +
            bdmSet.map(b => `<option value="${b}" ${b === selBDM ? 'selected' : ''}>${b}</option>`).join('');
        $('pdProduct').innerHTML = '<option value="">All Products</option>' +
            prodSet.map(p => `<option value="${p}" ${p === selProduct ? 'selected' : ''}>${p}</option>`).join('');

        // Filter product and OSG data
        let filtP = productData;
        if (selRBM) filtP = filtP.filter(r => r.rbm === selRBM);
        if (selBDM) filtP = filtP.filter(r => r.bdm === selBDM);
        if (selProduct) filtP = filtP.filter(r => r.product === selProduct);

        // Build invoice lookup from product data — only count OSG entries that match a product invoice
        const productInvoices = new Set();
        filtP.forEach(r => { if (r.invoice) productInvoices.add(r.invoice); });

        // Build invoice lookup from ALL OSG data (for missed customers section)
        const osgInvoices = new Set();
        osgData.forEach(r => { if (r.invoice) osgInvoices.add(r.invoice); });

        // Filter OSG: only entries whose invoice exists in product data
        let filtO = osgData.filter(r => r.invoice && productInvoices.has(r.invoice));
        if (selProduct) filtO = filtO.filter(r => r.product === selProduct);

        // Product stats: group by product category
        const pByProd = groupBy(filtP, 'product');
        const oByProd = groupBy(filtO, 'product');
        const productStats = Object.keys(pByProd).map(name => {
            const pRows = pByProd[name] || [];
            const oRows = oByProd[name] || [];
            const pQty = pRows.reduce((s, r) => s + r.qty, 0);
            const oQty = oRows.reduce((s, r) => s + r.qty, 0);
            const pRev = pRows.reduce((s, r) => s + r.soldPrice, 0);
            const oRev = oRows.reduce((s, r) => s + r.soldPrice, 0);
            const qtyConv = pQty > 0 ? (oQty / pQty) * 100 : 0;
            const valConv = pRev > 0 ? (oRev / pRev) * 100 : 0;
            return { name, pQty, oQty, pRev, oRev, qtyConv, valConv };
        });

        // KPI summary
        const totalPRev = filtP.reduce((s, r) => s + r.soldPrice, 0);
        const totalORev = filtO.reduce((s, r) => s + r.soldPrice, 0);
        const totalPQty = filtP.reduce((s, r) => s + r.qty, 0);
        const totalOQty = filtO.reduce((s, r) => s + r.qty, 0);
        const overallQtyConv = totalPQty > 0 ? (totalOQty / totalPQty) * 100 : 0;
        const overallValConv = totalPRev > 0 ? (totalORev / totalPRev) * 100 : 0;
        const uniqueProducts = productStats.length;

        $('pdKpiRow').innerHTML = `
            <div class="lc-kpi"><span class="lc-kpi-label">Products</span><span class="lc-kpi-val">${uniqueProducts}</span></div>
            <div class="lc-kpi"><span class="lc-kpi-label">Qty Conversion</span><span class="lc-kpi-val conversion-text">${overallQtyConv.toFixed(2)}%</span></div>
            <div class="lc-kpi"><span class="lc-kpi-label">Val Conversion</span><span class="lc-kpi-val conversion-text">${overallValConv.toFixed(2)}%</span></div>
            <div class="lc-kpi"><span class="lc-kpi-label">Product Revenue</span><span class="lc-kpi-val">${fmtShort(totalPRev)}</span></div>
            <div class="lc-kpi"><span class="lc-kpi-label">OSG Revenue</span><span class="lc-kpi-val">${fmtShort(totalORev)}</span></div>
            <div class="lc-kpi"><span class="lc-kpi-label">Product Qty</span><span class="lc-kpi-val">${formatNumber(totalPQty)}</span></div>
        `;

        // ---- Top Products by Revenue ----
        const topByRev = [...productStats].sort((a, b) => b.pRev - a.pRev).slice(0, 20);
        if (topByRev.length > 0) {
            let html = `<table class="data-table">
                <thead><tr>
                    <th>#</th><th>Product</th><th>Prod Qty</th><th>OSG Qty</th><th>Qty Conv%</th>
                    <th>Val Conv%</th><th>Prod Revenue</th><th>OSG Revenue</th>
                </tr></thead><tbody>`;
            topByRev.forEach((e, i) => {
                const rank = i + 1;
                const rankBadge = rank <= 3 ? `<span class="rank-badge rank-${rank}">${rank}</span>` : `<span class="rank-num">${rank}</span>`;
                const convCls = e.qtyConv === 0 ? 'loss-val' : (e.qtyConv < 5 ? 'conv-warn' : 'conv-val');
                html += `<tr>
                    <td class="number-cell">${rankBadge}</td>
                    <td><strong>${e.name}</strong></td>
                    <td class="number-cell">${formatNumber(e.pQty)}</td>
                    <td class="number-cell">${formatNumber(e.oQty)}</td>
                    <td class="number-cell ${convCls}">${e.qtyConv.toFixed(2)}%</td>
                    <td class="number-cell conv-val">${e.valConv.toFixed(2)}%</td>
                    <td class="number-cell">${fmtShort(e.pRev)}</td>
                    <td class="number-cell">${fmtShort(e.oRev)}</td>
                </tr>`;
            });
            html += '</tbody></table>';
            $('pdTopRevTable').innerHTML = html;
        } else {
            $('pdTopRevTable').innerHTML = noDataHTML('No product data available.');
        }

        // ---- Top Products by Conversion ----
        const topByConv = [...productStats].filter(p => p.pQty >= 3).sort((a, b) => b.qtyConv - a.qtyConv).slice(0, 20);
        if (topByConv.length > 0) {
            let html = `<table class="data-table">
                <thead><tr>
                    <th>#</th><th>Product</th><th>Prod Qty</th><th>OSG Qty</th><th>Qty Conv%</th>
                    <th>Val Conv%</th><th>Prod Revenue</th><th>OSG Revenue</th>
                </tr></thead><tbody>`;
            topByConv.forEach((e, i) => {
                const rank = i + 1;
                const rankBadge = rank <= 3 ? `<span class="rank-badge rank-${rank}">${rank}</span>` : `<span class="rank-num">${rank}</span>`;
                html += `<tr>
                    <td class="number-cell">${rankBadge}</td>
                    <td><strong>${e.name}</strong></td>
                    <td class="number-cell">${formatNumber(e.pQty)}</td>
                    <td class="number-cell">${formatNumber(e.oQty)}</td>
                    <td class="number-cell profit-val">${e.qtyConv.toFixed(2)}%</td>
                    <td class="number-cell profit-val">${e.valConv.toFixed(2)}%</td>
                    <td class="number-cell">${fmtShort(e.pRev)}</td>
                    <td class="number-cell">${fmtShort(e.oRev)}</td>
                </tr>`;
            });
            html += '</tbody></table>';
            $('pdTopConvTable').innerHTML = html;
        } else {
            $('pdTopConvTable').innerHTML = noDataHTML('No products with enough quantity.');
        }

    }

    // ---- CUSTOMERS WITHOUT OSG PAGE ----
    $('btnCORefresh').addEventListener('click', renderCustomersOSGPage);
    $('btnCOExport').addEventListener('click', exportCustomersOSGCSV);
    document.querySelector('[data-section="customers-osg-section"]').addEventListener('click', () => {
        // Load saved statuses from Firebase first, then render
        loadCoStatuses(() => setTimeout(renderCustomersOSGPage, 50));
    });
    ['coRBM', 'coBDM', 'coProduct', 'coBranch', 'coSort'].forEach(id => {
        $(id).addEventListener('change', () => {
            if (id === 'coSort') window.coSortMode = $('coSort').value;
            renderCustomersOSGPage();
        });
    });

    function loadCoStatuses(callback) {
        if (typeof firebase === 'undefined') { if (callback) callback(); return; }
        firebase.database().ref('customerStatus').once('value').then(snap => {
            const data = snap.val() || {};
            Object.keys(data).forEach(inv => {
                coStatusMap[inv] = data[inv];
            });
            if (callback) callback();
        }).catch(() => { if (callback) callback(); });
    }

    function saveCoStatus(inv) {
        if (typeof firebase === 'undefined') return;
        const status = coStatusMap[inv] || { callStatus: null, interest: null };
        firebase.database().ref('customerStatus/' + inv).set(status).catch(e => console.warn('[Firebase] Status save failed:', e));
    }

    function renderCustomersOSGPage() {
        let missedUnique = [];

        if (window.sharedMissedUnique) {
            // We are in shared view mode — apply client-side filters
            document.querySelectorAll('#customers-osg-section .lowconv-controls').forEach(el => el.style.display = 'none');

            // Init shared filter state
            if (!window.coSharedFilter) window.coSharedFilter = { branch: '', staff: '', product: '', callStatus: '', interest: '' };
            const sf = window.coSharedFilter;

            // Build unique values for filter dropdowns
            const allData = window.sharedMissedUnique;
            const branches = [...new Set(allData.map(r => r.branch).filter(Boolean))].sort();
            const staffs   = [...new Set(allData.map(r => r.staff).filter(Boolean))].sort();
            const products = [...new Set(allData.map(r => r.product).filter(Boolean))].sort();

            // Inject shared filter bar into page (only once)
            let filterBarEl = $('coSharedFilterBar');
            if (!filterBarEl) {
                const bar = document.createElement('div');
                bar.id = 'coSharedFilterBar';
                bar.style.cssText = 'display:flex;gap:10px;flex-wrap:wrap;margin-bottom:16px;padding:14px 18px;background:var(--bg-card);border:1px solid var(--border);border-radius:12px;align-items:flex-end;';
                $('coMissedTable').parentElement.insertBefore(bar, $('coMissedTable'));
                filterBarEl = bar;
            }

            const sel = (id, label, opts, val) => `
                <div style="display:flex;flex-direction:column;gap:4px;min-width:130px;flex:1;">
                    <label style="font-size:0.72rem;font-weight:600;color:var(--text-muted);text-transform:uppercase;letter-spacing:0.05em;">${label}</label>
                    <select id="${id}" onchange="window.coSharedFilterChange('${id}',this.value)" style="padding:8px 10px;background:var(--bg-input);border:1px solid var(--border);border-radius:8px;color:var(--text-primary);font-family:inherit;font-size:0.85rem;">
                        <option value="">All</option>
                        ${opts.map(o => `<option value="${o}" ${val===o?'selected':''}>${o}</option>`).join('')}
                    </select>
                </div>`;

            filterBarEl.innerHTML =
                sel('coSF_branch',  'Branch',      branches, sf.branch) +
                sel('coSF_staff',   'Staff',       staffs,   sf.staff) +
                sel('coSF_product', 'Product',     products, sf.product) +
                sel('coSF_call',    'Call Status', ['connected','disconnected'], sf.callStatus) +
                sel('coSF_int',     'Interest',    ['interested','not-interested'], sf.interest) +
                `<button onclick="window.coSharedFilterReset()" style="padding:8px 14px;border-radius:8px;border:1px solid var(--border);background:transparent;color:var(--text-muted);font-family:inherit;cursor:pointer;font-size:0.82rem;white-space:nowrap;">✕ Clear</button>`;

            window.coSharedFilterChange = (id, val) => {
                const map = { coSF_branch: 'branch', coSF_staff: 'staff', coSF_product: 'product', coSF_call: 'callStatus', coSF_int: 'interest' };
                window.coSharedFilter[map[id]] = val;
                renderCustomersOSGPage();
            };
            window.coSharedFilterReset = () => {
                window.coSharedFilter = { branch: '', staff: '', product: '', callStatus: '', interest: '' };
                renderCustomersOSGPage();
            };

            // Apply filters to data
            missedUnique = allData.filter(r => {
                if (sf.branch  && r.branch  !== sf.branch)  return false;
                if (sf.staff   && r.staff   !== sf.staff)   return false;
                if (sf.product && r.product !== sf.product) return false;
                if (sf.callStatus) {
                    const inv = r.invoice || '';
                    const st = (coStatusMap[inv] || {}).callStatus;
                    if (st !== sf.callStatus) return false;
                }
                if (sf.interest) {
                    const inv = r.invoice || '';
                    const st = (coStatusMap[inv] || {}).interest;
                    if (st !== sf.interest) return false;
                }
                return true;
            });

            $('coMissedCount').textContent = `${missedUnique.length} of ${allData.length} customers (Shared View)`;

        } else {
            // Standard dynamic processing
            if (productData.length === 0) {
                $('coMissedTable').innerHTML = noDataHTML('Upload data and generate reports first.');
                $('coMissedCount').textContent = '0 customers';
                $('coMissedCount').style.background = '';
                return;
            }

            const selRBM = $('coRBM').value;
            const selBDM = $('coBDM').value;
            const selProduct = $('coProduct').value;
            const selBranch = $('coBranch').value;

            // Populate filter dropdowns (preserve selection)
            const rbmSet = [...new Set(productData.map(r => r.rbm).filter(Boolean))].sort();
            const bdmSet = [...new Set(productData.map(r => r.bdm).filter(Boolean))].sort();
            const prodSet = [...new Set(productData.map(r => r.product).filter(Boolean))].sort();
            const branchSet = [...new Set(productData.map(r => r.branch).filter(Boolean))].sort();

            $('coRBM').innerHTML = '<option value="">All RBMs</option>' + rbmSet.map(r => `<option value="${r}" ${r === selRBM ? 'selected' : ''}>${r}</option>`).join('');
            $('coBDM').innerHTML = '<option value="">All BDMs</option>' + bdmSet.map(b => `<option value="${b}" ${b === selBDM ? 'selected' : ''}>${b}</option>`).join('');
            $('coProduct').innerHTML = '<option value="">All Products</option>' + prodSet.map(p => `<option value="${p}" ${p === selProduct ? 'selected' : ''}>${p}</option>`).join('');
            $('coBranch').innerHTML = '<option value="">All Branches</option>' + branchSet.map(b => `<option value="${b}" ${b === selBranch ? 'selected' : ''}>${b}</option>`).join('');

            // Filter product rows
            let filtP = productData;
            if (selRBM) filtP = filtP.filter(r => r.rbm === selRBM);
            if (selBDM) filtP = filtP.filter(r => r.bdm === selBDM);
            if (selProduct) filtP = filtP.filter(r => r.product === selProduct);
            if (selBranch) filtP = filtP.filter(r => r.branch === selBranch);

            // Build OSG invoice set
            const osgInvoices = new Set();
            osgData.forEach(r => { if (r.invoice) osgInvoices.add(r.invoice); });

            // Find product rows with no matching OSG invoice, deduplicated by invoice
            const seenInv = new Set();
            filtP.forEach(r => {
                if (r.invoice && !osgInvoices.has(r.invoice) && !seenInv.has(r.invoice)) {
                    seenInv.add(r.invoice);
                    missedUnique.push(r);
                }
            });

            $('coMissedCount').textContent = `${missedUnique.length} customers`;
        }

        // Sort based on user selection
        const sortMode = (typeof window.coSortMode !== 'undefined') ? window.coSortMode : 'value-desc';
        if (sortMode === 'name-asc') {
            missedUnique.sort((a, b) => (a.customerName || '').localeCompare(b.customerName || ''));
        } else if (sortMode === 'name-desc') {
            missedUnique.sort((a, b) => (b.customerName || '').localeCompare(a.customerName || ''));
        } else if (sortMode === 'value-asc') {
            missedUnique.sort((a, b) => (a.soldPrice || 0) - (b.soldPrice || 0));
        } else {
            missedUnique.sort((a, b) => (b.soldPrice || 0) - (a.soldPrice || 0));
        }

        if (missedUnique.length === 0) {
            $('coMissedTable').innerHTML = noDataHTML('All invoices have OSG entries — great conversion! 🎉');
            return;
        }

        // Store filtered rows globally for pagination and single-row updates
        coCurrentRows = missedUnique;
        coDisplayLimit = 100;

        // ---- Stats Bar (with IDs for in-place update) ----
        const total        = missedUnique.length;
        const connected    = missedUnique.filter(r => (coStatusMap[r.invoice||'']||{}).callStatus === 'connected').length;
        const disconnected = missedUnique.filter(r => (coStatusMap[r.invoice||'']||{}).callStatus === 'disconnected').length;
        const interested   = missedUnique.filter(r => (coStatusMap[r.invoice||'']||{}).interest === 'interested').length;
        const notInterested= missedUnique.filter(r => (coStatusMap[r.invoice||'']||{}).interest === 'not-interested').length;
        const notCalled    = total - connected - disconnected;

        const statsBar = `
        <div id="coStatsBar" class="kpi-grid" style="margin-bottom:24px;">
            <div class="kpi-card" style="border-top: 4px solid var(--text-muted); height: 100px; padding: 16px 20px;">
                <div class="kpi-content">
                    <div class="kpi-label">Total Results</div>
                    <div id="coStat_total" class="kpi-value">${total}</div>
                </div>
            </div>
            <div class="kpi-card" style="border-top: 4px solid var(--text-secondary); height: 100px; padding: 16px 20px;">
                <div class="kpi-content">
                    <div class="kpi-label">Not Called</div>
                    <div id="coStat_notCalled" class="kpi-value" style="color:var(--text-secondary);">${notCalled}</div>
                </div>
            </div>
            <div class="kpi-card" style="border-top: 4px solid var(--accent-blue); height: 100px; padding: 16px 20px;">
                <div class="kpi-content">
                    <div class="kpi-label">Connected</div>
                    <div id="coStat_connected" class="kpi-value" style="color:var(--accent-blue);">${connected}</div>
                </div>
            </div>
            <div class="kpi-card" style="border-top: 4px solid var(--accent-purple); height: 100px; padding: 16px 20px;">
                <div class="kpi-content">
                    <div class="kpi-label">Disconnected</div>
                    <div id="coStat_disconnected" class="kpi-value" style="color:var(--accent-purple);">${disconnected}</div>
                </div>
            </div>
            <div class="kpi-card" style="border-top: 4px solid var(--accent-green); height: 100px; padding: 16px 20px;">
                <div class="kpi-content">
                    <div class="kpi-label">Interested</div>
                    <div id="coStat_interested" class="kpi-value" style="color:var(--accent-green);">${interested}</div>
                </div>
            </div>
            <div class="kpi-card" style="border-top: 4px solid var(--accent-red); height: 100px; padding: 16px 20px;">
                <div class="kpi-content">
                    <div class="kpi-label">Not Interested</div>
                    <div id="coStat_notInterested" class="kpi-value" style="color:var(--accent-red);">${notInterested}</div>
                </div>
            </div>
        </div>`;

        // ---- Caller Selector ----
        const callerSelector = `
        <div class="filters-bar" style="padding: 16px 24px; margin-bottom:24px;">
            <div class="filter-group" style="min-width: auto; flex: 0;">
                <label>👤 Session Caller</label>
                <div style="display:flex;gap:10px;margin-top:4px;">
                    ${CO_CALLERS.map(c => `
                    <button onclick="window.selectCoCaller('${c.name}')" class="nav-item ${currentCaller===c.name ? 'active' : ''}" style="
                        padding:8px 16px; border-radius:30px; font-size:0.88rem; cursor:pointer;
                        ${currentCaller===c.name ? 'background:'+c.bg+'; color:'+c.color+'; border-color:'+c.color : ''}
                    ">
                        <span style="width:20px;height:20px;border-radius:50%;background:${c.color};color:#fff;display:inline-flex;align-items:center;justify-content:center;font-size:0.65rem;margin-right:4px;">${c.name[0]}</span>
                        ${c.name}
                    </button>`).join('')}
                </div>
            </div>
            <div style="margin-left:auto; display:flex; flex-direction:column; align-items:flex-end;">
                ${currentCaller
                    ? `<span style="font-size:0.85rem;color:var(--text-secondary);">Logged in as <strong style="color:var(--accent-blue);">${currentCaller}</strong></span>`
                    : '<div class="status-badge" style="background:rgba(245,158,11,0.15); color:var(--accent-amber);">⚠️ Please select your name</div>'}
            </div>
        </div>`;

        // ---- Table scaffold + first page of rows ----
        const tableHeader = `<div style="overflow-x:auto;"><table id="coMainTable" class="data-table" style="width:100%;border-collapse:collapse;font-size:0.875rem;">
            <thead><tr style="background:var(--bg-input);border-bottom:2px solid var(--border);">
                <th style="padding:12px 10px;text-align:left;font-weight:600;color:var(--text-muted);font-size:0.78rem;letter-spacing:0.05em;text-transform:uppercase;white-space:nowrap;">#</th>
                <th style="padding:12px 10px;text-align:left;font-weight:600;color:var(--text-muted);font-size:0.78rem;letter-spacing:0.05em;text-transform:uppercase;white-space:nowrap;">Invoice</th>
                <th style="padding:12px 10px;text-align:left;font-weight:600;color:var(--text-muted);font-size:0.78rem;letter-spacing:0.05em;text-transform:uppercase;white-space:nowrap;">Customer</th>
                <th style="padding:12px 10px;text-align:left;font-weight:600;color:var(--text-muted);font-size:0.78rem;letter-spacing:0.05em;text-transform:uppercase;white-space:nowrap;">Interest</th>
                <th style="padding:12px 10px;text-align:left;font-weight:600;color:var(--text-muted);font-size:0.78rem;letter-spacing:0.05em;text-transform:uppercase;white-space:nowrap;">Contact &amp; Call Status</th>
                <th style="padding:12px 10px;text-align:left;font-weight:600;color:var(--text-muted);font-size:0.78rem;letter-spacing:0.05em;text-transform:uppercase;white-space:nowrap;">Remarks</th>
                <th style="padding:12px 10px;text-align:left;font-weight:600;color:var(--text-muted);font-size:0.78rem;letter-spacing:0.05em;text-transform:uppercase;white-space:nowrap;">Branch</th>
                <th style="padding:12px 10px;text-align:left;font-weight:600;color:var(--text-muted);font-size:0.78rem;letter-spacing:0.05em;text-transform:uppercase;white-space:nowrap;">Product</th>
                <th style="padding:12px 10px;text-align:right;font-weight:600;color:var(--text-muted);font-size:0.78rem;letter-spacing:0.05em;text-transform:uppercase;white-space:nowrap;">Value</th>
            </tr></thead><tbody id="coTableBody">`;

        let rowsHTML = '';
        missedUnique.slice(0, coDisplayLimit).forEach((r, i) => {
            rowsHTML += buildCoRowHTML(r, i);
        });

        const remaining = missedUnique.length - coDisplayLimit;
        const loadMoreBtn = remaining > 0
            ? `<div id="coLoadMoreWrap" style="text-align:center;margin-top:24px;">
                <button onclick="window.coLoadMore()" class="btn-apply-filter" style="margin: 0 auto;">
                    ↧ Load ${Math.min(remaining, 100)} more (${remaining} remaining)
                </button>
               </div>` : '';

        $('coMissedTable').innerHTML = statsBar + callerSelector + tableHeader + rowsHTML + '</tbody></table></div>' + loadMoreBtn;

        // ---- Handlers ----
        window.selectCoCaller = function(name) {
            currentCaller = name;
            localStorage.setItem('co_caller', name);
            renderCustomersOSGPage();
        };

        window.coLoadMore = function() {
            const tbody = $('coTableBody');
            if (!tbody) return;
            const from = coDisplayLimit;
            coDisplayLimit = Math.min(coDisplayLimit + 100, coCurrentRows.length);
            let extra = '';
            coCurrentRows.slice(from, coDisplayLimit).forEach((r, i) => {
                extra += buildCoRowHTML(r, from + i);
            });
            tbody.insertAdjacentHTML('beforeend', extra);
            const remaining2 = coCurrentRows.length - coDisplayLimit;
            const wrap = $('coLoadMoreWrap');
            if (wrap) {
                if (remaining2 > 0) {
                    wrap.innerHTML = `<button onclick="window.coLoadMore()" style="
                        padding:10px 28px;border-radius:20px;border:1.5px solid var(--accent);
                        background:transparent;color:var(--accent);font-family:inherit;
                        font-size:0.88rem;font-weight:600;cursor:pointer;">↧ Load ${Math.min(remaining2,100)} more (${remaining2} remaining)</button>`;
                } else {
                    wrap.remove();
                }
            }
        };

        window.toggleCoCall = function(inv, status) {
            if (!currentCaller && status !== "") {
                alert('Please select your name (Harmiya / Aswathi / Shikha) at the top of the page before logging a call.');
                // Revert UI to match state since we rejected it
                updateCoSingleRow(inv);
                return;
            }
            if (!coStatusMap[inv]) coStatusMap[inv] = { callStatus: null, interest: null, calledBy: null, remarks: '' };
            coStatusMap[inv].callStatus = status === "" ? null : status;
            coStatusMap[inv].calledBy   = coStatusMap[inv].callStatus ? currentCaller : null;
            saveCoStatus(inv);
            updateCoSingleRow(inv);
            updateCoStatsInPlace();
        };

        window.toggleCoInterest = function(inv, status) {
            if (!coStatusMap[inv]) coStatusMap[inv] = { callStatus: null, interest: null, calledBy: null, remarks: '' };
            coStatusMap[inv].interest = status === "" ? null : status;
            saveCoStatus(inv);
            updateCoSingleRow(inv);
            updateCoStatsInPlace();
        };

        window.saveCoRemark = function(inv, remark) {
            if (!coStatusMap[inv]) coStatusMap[inv] = { callStatus: null, interest: null, calledBy: null, remarks: '' };
            coStatusMap[inv].remarks = remark;
            saveCoStatus(inv);
        };
    }

    // Updates just the 6 stat numbers in-place (no table re-render)
    function updateCoStatsInPlace() {
        const rows = coCurrentRows;
        const connected    = rows.filter(r => (coStatusMap[r.invoice||'']||{}).callStatus === 'connected').length;
        const disconnected = rows.filter(r => (coStatusMap[r.invoice||'']||{}).callStatus === 'disconnected').length;
        const interested   = rows.filter(r => (coStatusMap[r.invoice||'']||{}).interest === 'interested').length;
        const notInterested= rows.filter(r => (coStatusMap[r.invoice||'']||{}).interest === 'not-interested').length;
        const notCalled    = rows.length - connected - disconnected;
        const set = (id, val) => { const el = $(id); if (el) el.textContent = val; };
        set('coStat_total', rows.length);
        set('coStat_notCalled', notCalled);
        set('coStat_connected', connected);
        set('coStat_disconnected', disconnected);
        set('coStat_interested', interested);
        set('coStat_notInterested', notInterested);
    }

    // Replaces a single row in-place without touching the rest of the table
    // Replaces a single row in-place without touching the rest of the table
    function updateCoSingleRow(inv) {
        const rowEl = document.getElementById('co-row-' + inv);
        if (!rowEl) return;
        const idx = coCurrentRows.findIndex((r, i) => (r.invoice || String(i)) === inv);
        if (idx === -1) return;
        const newRowHTML = buildCoRowHTML(coCurrentRows[idx], idx);
        
        // Use a temporary container to swap the element while preserving transition
        const temp = document.createElement('tbody');
        temp.innerHTML = newRowHTML;
        const newRow = temp.firstElementChild;
        
        newRow.style.background = 'rgba(59, 130, 246, 0.15)';
        rowEl.replaceWith(newRow);
        
        setTimeout(() => {
            const el = document.getElementById('co-row-' + inv);
            if (el) el.style.background = '';
        }, 800);
    }

    // Builds HTML for a single table row
    function buildCoRowHTML(r, i) {
        const inv = r.invoice || String(i);
        const st  = coStatusMap[inv] || { callStatus: null, interest: null, calledBy: null, remarks: '' };

        const rowClass = st.interest === 'interested' ? 'insight-success' 
                       : st.interest === 'not-interested' ? 'insight-danger' 
                       : '';

        const interestBtn = `
            <select onchange="window.toggleCoInterest('${inv}', this.value)" class="crm-select" style="border-color:${st.interest === 'interested' ? 'var(--accent-green)' : (st.interest === 'not-interested' ? 'var(--accent-red)' : 'var(--border)')};">
                <option value="" ${!st.interest ? 'selected' : ''}>- Select -</option>
                <option value="interested" ${st.interest === 'interested' ? 'selected' : ''}>✅ Interested</option>
                <option value="not-interested" ${st.interest === 'not-interested' ? 'selected' : ''}>❌ Not Interested</option>
            </select>`;

        const callerInfo = st.calledBy ? (() => {
            const c = CO_CALLERS.find(x => x.name === st.calledBy) || { color:'#94a3b8', bg:'rgba(148, 163, 184, 0.1)' };
            return `<div style="display:inline-flex;align-items:center;gap:6px;padding:4px 10px;border-radius:12px;background:${c.bg};color:${c.color};font-size:0.75rem;font-weight:700;margin-top:6px;border:1px solid ${c.color}22;">
                <span style="width:16px;height:16px;border-radius:50%;background:${c.color};color:#fff;display:inline-flex;align-items:center;justify-content:center;font-size:0.6rem;font-weight:800;">${st.calledBy[0]}</span>
                ${st.calledBy}</div>`;
        })() : '';

        const callBtns = r.customerNo ? `
            <div style="display:flex;flex-direction:column;gap:6px;margin-top:8px;align-items:flex-start;">
                <select onchange="window.toggleCoCall('${inv}', this.value)" class="crm-select" style="border-color:${st.callStatus === 'connected' ? 'var(--accent-blue)' : (st.callStatus === 'disconnected' ? 'var(--accent-purple)' : 'var(--border)')};">
                    <option value="" ${!st.callStatus ? 'selected' : ''}>- Status -</option>
                    <option value="connected" ${st.callStatus === 'connected' ? 'selected' : ''}>📞 Connected</option>
                    <option value="disconnected" ${st.callStatus === 'disconnected' ? 'selected' : ''}>📵 Disconnected</option>
                </select>
                ${callerInfo}
            </div>` : '';

        return `<tr id="co-row-${inv}" class="${rowClass}">
            <td><span class="text-muted">${i+1}</span></td>
            <td><code style="color:var(--text-secondary); opacity:0.8;">${r.invoice}</code></td>
            <td><div style="font-weight:700; color:var(--text-primary); font-size:0.95rem;">${r.customerName||'—'}</div></td>
            <td>${interestBtn}</td>
            <td>
                <div style="display:flex;align-items:center;gap:8px;margin-bottom:4px;">
                    <span style="font-weight:600;color:var(--text-primary);">${r.customerNo||'—'}</span>
                    ${r.customerNo?`<a href="tel:${r.customerNo}" class="status-badge" style="background:rgba(59,130,246,0.1); color:var(--accent-blue); padding:4px;" title="Call"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72 12.84 12.84 0 0 0 .7 2.81 2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45 12.84 12.84 0 0 0 2.81.7A2 2 0 0 1 22 16.92z"/></svg></a>`:''}
                    ${r.customerNo?`<a href="https://wa.me/91${r.customerNo.replace(/\D/g,'')}" target="_blank" class="status-badge" style="background:rgba(37,211,102,0.1); color:#25D366; padding:4px;" title="WhatsApp"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 11.5a8.38 8.38 0 0 1-.9 3.8 8.5 8.5 0 0 1-7.6 4.7 8.38 8.38 0 0 1-3.8-.9L3 21l1.9-5.7a8.38 8.38 0 0 1-.9-3.8 8.5 8.5 0 0 1 4.7-7.6 8.38 8.38 0 0 1 3.8-.9h.5a8.48 8.48 0 0 1 8 8v.5z"/></svg></a>`:''}
                </div>
                ${callBtns}
            </td>
            <td><input type="text" value="${q(st.remarks || '')}" onchange="window.saveCoRemark('${inv}', this.value)" placeholder="Add remarks..." class="crm-input" style="width:160px;" /></td>
            <td><span class="text-muted" style="font-size:0.85rem;">${r.branch||'—'}</span></td>
            <td><span class="text-secondary" style="font-size:0.85rem; font-weight:500;">${r.product||'—'}</span></td>
            <td style="text-align:right;"><div class="conv-warn" style="font-size:1rem;">${fmtShort(r.soldPrice)}</div></td>
        </tr>`;
    }

    function exportCustomersOSGCSV() {
        if (productData.length === 0) return;
        const selRBM = $('coRBM').value;
        const selBDM = $('coBDM').value;
        const selProduct = $('coProduct').value;
        const selBranch = $('coBranch').value;

        let filtP = productData;
        if (selRBM) filtP = filtP.filter(r => r.rbm === selRBM);
        if (selBDM) filtP = filtP.filter(r => r.bdm === selBDM);
        if (selProduct) filtP = filtP.filter(r => r.product === selProduct);
        if (selBranch) filtP = filtP.filter(r => r.branch === selBranch);

        const osgInvoices = new Set();
        osgData.forEach(r => { if (r.invoice) osgInvoices.add(r.invoice); });

        const seenInv = new Set();
        const missedUnique = [];
        filtP.forEach(r => {
            if (r.invoice && !osgInvoices.has(r.invoice) && !seenInv.has(r.invoice)) {
                seenInv.add(r.invoice);
                missedUnique.push(r);
            }
        });

        if (missedUnique.length === 0) return;
        const hdr = ['#', 'Invoice No', 'Customer Name', 'Customer No', 'Staff', 'Branch', 'RBM', 'BDM', 'Product', 'Qty', 'Sold Price'];
        const lines = [hdr.join(',')];
        missedUnique.forEach((r, i) => {
            lines.push([i + 1, q(r.invoice), q(r.customerName || ''), q(r.customerNo || ''),
            q(r.staff || ''), q(r.branch || ''), q(r.rbm || ''), q(r.bdm || ''),
            q(r.product || ''), r.qty, r.soldPrice.toFixed(0)].join(','));
        });
        downloadCSV(lines.join('\n'), 'customers_without_osg.csv');
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

    // ---- DEEP INSIGHTS PAGE ----
    $('btnRefreshInsights').addEventListener('click', renderDeepInsightsPage);
    document.querySelector('[data-section="insights-section"]').addEventListener('click', () => {
        setTimeout(renderDeepInsightsPage, 50);
    });

    function renderDeepInsightsPage() {
        if (productData.length === 0) {
            $('insightsContentWrapper').innerHTML = noDataHTML('Upload data and generate reports first.');
            return;
        }

        const wrapper = $('insightsContentWrapper');
        wrapper.innerHTML = '<div style="text-align:center; padding: 40px; color:var(--text-muted);"><div class="spinner" style="margin: 0 auto 16px;"></div><p>Analyzing datasets with heuristics...</p></div>';

        // Simulate an "AI Analysis" delay for UX impact
        setTimeout(() => {
            try {
                const html = generateDeepInsightsHTML();
                wrapper.innerHTML = html;
            } catch (err) {
                console.error(err);
                wrapper.innerHTML = noDataHTML('Error generating insights: ' + err.message);
            }
        }, 800);
    }

    function generateDeepInsightsHTML() {
        // Gather key metrics for insights
        const pG_branch = groupBy(filteredProduct, 'branch');
        const oG_branch = groupBy(filteredOSG, 'branch');
        const pG_staff = groupBy(filteredProduct, 'staff');
        const oG_staff = groupBy(filteredOSG, 'staff');

        // 1. Worst Performing Branches (High Qty, near 0% OSG)
        const branchStats = Object.keys(pG_branch).map(b => {
            const pQ = pG_branch[b].reduce((s, r) => s + r.qty, 0);
            const oQ = (oG_branch[b] || []).reduce((s, r) => s + r.qty, 0);
            const valP = pG_branch[b].reduce((s, r) => s + r.soldPrice, 0);
            return { branch: b, pQ, oQ, valP, conv: pQ > 0 ? (oQ / pQ) * 100 : 0 };
        }).filter(b => b.pQ > 5 && b.branch !== 'Unknown').sort((a, b) => a.conv - b.conv);

        const worstBranches = branchStats.slice(0, 3);

        // 2. Missed Premium Opportunities (Products > 50,000 sold without OSG)
        const osgInvoices = new Set(osgData.map(r => r.invoice).filter(Boolean));
        const premiumMisses = filteredProduct.filter(r =>
            r.soldPrice > 50000 && r.invoice && !osgInvoices.has(r.invoice)
        ).sort((a, b) => b.soldPrice - a.soldPrice).slice(0, 5);

        // 3. Bottom Staff By Attachment Rate (Min 10 sales)
        const staffStats = Object.keys(pG_staff).map(s => {
            const pQ = pG_staff[s].reduce((sum, r) => sum + r.qty, 0);
            const oQ = (oG_staff[s] || []).reduce((sum, r) => sum + r.qty, 0);
            return { staff: s, pQ, oQ, conv: pQ > 0 ? (oQ / pQ) * 100 : 0, branch: pG_staff[s][0].branch };
        }).filter(s => s.pQ >= 10 && s.staff !== 'Unknown').sort((a, b) => a.conv - b.conv);

        const worstStaff = staffStats.slice(0, 3);

        let html = '<div style="display:flex; flex-direction:column; gap:24px;">';

        // SECTION 1: Critical Focus Areas
        html += '<div class="conversion-card" style="border-left: 4px solid var(--loss); border-radius: 8px; padding: 20px; background:var(--bg-card);">';
        html += '<h3 style="margin-top:0; color:var(--loss); display:flex; align-items:center; gap:8px;"><svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg> Critical Focus Areas</h3>';

        html += '<div style="margin-bottom: 16px;"><strong>🚨 High Volume, Low Conversion Branches:</strong><ul style="margin:8px 0 0 20px; color:var(--text-muted); line-height: 1.6;">';
        if (worstBranches.length) worstBranches.forEach(b => html += `<li><strong>${b.branch}</strong>: ${b.pQ} products sold but only ${b.oQ} OSG attached (${b.conv.toFixed(1)}%). Estimated missed tracking revenue: ${fmtShort(b.valP * 0.1)}</li>`);
        else html += '<li>No significantly underperforming branches detected.</li>';
        html += '</ul></div>';

        html += '<div style="margin-bottom: 16px;"><strong>💎 Missed Premium Device Attachments:</strong><ul style="margin:8px 0 0 20px; color:var(--text-muted); line-height: 1.6;">';
        if (premiumMisses.length) premiumMisses.forEach(m => html += `<li><strong>${m.staff} (${m.branch})</strong> sold a ${m.product} for ${fmtShort(m.soldPrice)} without OSG (Inv: ${m.invoice}).</li>`);
        else html += '<li>Great job! High-value premium products seem to be attached correctly.</li>';
        html += '</ul></div>';

        html += '<div><strong>👤 Highest Opportunity Staff:</strong><ul style="margin:8px 0 0 20px; color:var(--text-muted); line-height: 1.6;">';
        if (worstStaff.length) worstStaff.forEach(s => html += `<li><strong>${s.staff} (${s.branch})</strong>: Delivered ${s.pQ} units physically but achieved only ${s.conv.toFixed(1)}% conversion.</li>`);
        else html += '<li>Staff metrics look solid across the board (or volume threshold not met).</li>';
        html += '</ul></div>';

        html += '</div>';

        // SECTION 2: Automated Root Cause Analysis
        let reasonHTML = '';
        if (worstBranches.length && worstBranches[0].conv < 5) {
            reasonHTML += `<li><strong>Systemic Branch Failure (${worstBranches[0].branch}):</strong> Conversion is near zero (${worstBranches[0].conv.toFixed(1)}%) despite moving ${worstBranches[0].pQ} units. This indicates a store-wide knowledge gap or a leadership failure to enforce pitching at the POS, rather than individual poor performance.</li>`;
        }
        if (premiumMisses.length >= 2) {
            reasonHTML += `<li><strong>Premium Pitch Avoidance:</strong> Found multiple premium devices > ₹50K sold with no OSG attached. Sales reps might be avoiding the OSG pitch on high-ticket items out of fear of losing the primary sale due to total cart value shock.</li>`;
        }
        const topPerformers = staffStats.filter(s => s.conv > 20);
        if (topPerformers.length > 0 && worstStaff.length > 0) {
            reasonHTML += `<li><strong>Inconsistent Training Deployment:</strong> The gap between top performers (e.g. ${topPerformers[topPerformers.length - 1].staff} at ${topPerformers[topPerformers.length - 1].conv.toFixed(1)}%) and bottom performers (${worstStaff[0].conv.toFixed(1)}%) is massive. Core OSG knowledge is siloed to specific top individuals.</li>`;
        }

        if (reasonHTML) {
            html += '<div class="conversion-card" style="border-left: 4px solid var(--accent); border-radius: 8px; padding: 20px; background:var(--bg-card);">';
            html += '<h3 style="margin-top:0; color:var(--accent); display:flex; align-items:center; gap:8px;"><svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="16" x2="12" y2="12"/><line x1="12" y1="8" x2="12.01" y2="8"/></svg> Automated Root Cause Analysis</h3>';
            html += `<ul style="margin:8px 0 0 20px; color:var(--text-muted); line-height:1.6;">${reasonHTML}</ul>`;
            html += '</div>';
        }

        // SECTION 3: Action Plan
        html += '<div class="conversion-card" style="border-left: 4px solid var(--primary); border-radius: 8px; padding: 20px; background:var(--bg-card);">';
        html += '<h3 style="margin-top:0; color:var(--primary); display:flex; align-items:center; gap:8px;"><svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg> Immediate Action Plan</h3>';
        html += '<ol style="margin:8px 0 0 24px; color:var(--text-muted); line-height:1.8;">';
        if (worstBranches.length) html += `<li><strong>RBM/BDM Intercept:</strong> Immediately deploy regional trainers to ${worstBranches.map(b => `<strong>${b.branch}</strong>`).join(', ')} for POS floor shadowing.</li>`;
        if (premiumMisses.length) html += `<li><strong>Premium Bundling Rule:</strong> Institute a strict rule that any manager override/discount on products over ₹50K ideally requires an OSG attachment commitment.</li>`;
        if (worstStaff.length) html += `<li><strong>Targeted PIPs:</strong> Place <strong>${worstStaff.map(s => `${s.staff}`).join(', ')}</strong> on an accelerated 7-day OSG pitch improvement plan.</li>`;
        html += '<li><strong>Daily Morning Brief:</strong> Have branch managers physically review the "Customers Without OSG" dashboard list from yesterday\'s data before the store opens to identify missed pitch opportunities and contact customers via the WhatsApp quick-links.</li>';
        html += '</ol>';
        html += '</div>';

        html += '</div>'; // End flex wrapper
        return html;
    }

})();

