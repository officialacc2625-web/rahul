// ============================================================
// Analytics Portal ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â Conversion Reports
// Dual-file: Product Data + OSG Data
// Value Conversion = OSG Sold Price / Product Sold Price
// Qty Conversion   = OSG Quantity  / Product Quantity
// ============================================================

(function () {
    'use strict';
    window.onerror = function (msg, url, lineNo, columnNo, error) {
        console.error('Global Error: ' + msg + ' at ' + lineNo + ':' + columnNo);
        return false;
    };

    // ---- STATE ----
    console.log('[app.js] Loaded successfully');

    let productData = [];      // Parsed rows from Product file
    let osgData = [];          // Parsed rows from OSG file
    let amcData = [];          // Parsed rows from LG AMC file (optional)
    let samsungData = [];      // Parsed rows from Samsung file (optional)
    let allData = [];          // Product + AMC merged (for profit/loss)
    let filteredProduct = [];  // After filters
    let filteredOSG = [];      // After filters
    let filteredAMC = [];      // After filters (AMC)
    let filteredSamsung = [];  // After filters (Samsung)
    let filteredAll = [];      // After filters (product+amc)
    let chartInstances = {};
    let pincodeAreaCache = {}; // Cache for pincode area names API

    // ---- INDEXEDDB FOR MONTHLY CACHING ----
    const DB_NAME = 'myGAnalyticsDB';
    const DB_VERSION = 1;
    const STORE_NAME = 'monthlyData';

    function initDB() {
        return new Promise((resolve, reject) => {
            const request = indexedDB.open(DB_NAME, DB_VERSION);
            request.onupgradeneeded = (e) => {
                const db = e.target.result;
                if (!db.objectStoreNames.contains(STORE_NAME)) {
                    db.createObjectStore(STORE_NAME, { keyPath: 'month' });
                }
            };
            request.onsuccess = () => resolve(request.result);
            request.onerror = () => reject(request.error);
        });
    }

    async function saveMonthlyDataToDB(month, data) {
        try {
            const db = await initDB();
            const tx = db.transaction(STORE_NAME, 'readwrite');
            const store = tx.objectStore(STORE_NAME);
            store.put({ month, ...data });
            return new Promise((resolve, reject) => {
                tx.oncomplete = resolve;
                tx.onerror = () => reject(tx.error);
            });
        } catch (e) {
            console.error('[IndexedDB] Failed to save data', e);
        }
    }

    async function loadMonthlyDataFromDB(month) {
        try {
            const db = await initDB();
            const tx = db.transaction(STORE_NAME, 'readonly');
            const store = tx.objectStore(STORE_NAME);
            const request = store.get(month);
            return new Promise((resolve, reject) => {
                request.onsuccess = () => resolve(request.result);
                request.onerror = () => reject(request.error);
            });
        } catch (e) {
            console.error('[IndexedDB] Failed to load data', e);
            return null;
        }
    }

    async function getAllMonthsFromDB() {
        try {
            const db = await initDB();
            const tx = db.transaction(STORE_NAME, 'readonly');
            const store = tx.objectStore(STORE_NAME);
            const request = store.getAllKeys();
            return new Promise((resolve, reject) => {
                request.onsuccess = () => resolve(request.result || []);
                request.onerror = () => reject(request.error);
            });
        } catch (e) {
            console.error('[IndexedDB] Failed to get months', e);
            return [];
        }
    }

    async function getAllMonthlyDataFromDB() {
        try {
            const db = await initDB();
            const tx = db.transaction(STORE_NAME, 'readonly');
            const store = tx.objectStore(STORE_NAME);
            const request = store.getAll();
            return new Promise((resolve, reject) => {
                request.onsuccess = () => resolve(request.result || []);
                request.onerror = () => reject(request.error);
            });
        } catch (e) {
            console.error('[IndexedDB] Failed to get all data', e);
            return [];
        }
    }

    async function clearAllMonthlyDataFromDB() {
        try {
            const db = await initDB();
            const tx = db.transaction(STORE_NAME, 'readwrite');
            const store = tx.objectStore(STORE_NAME);
            store.clear();
            return new Promise((resolve, reject) => {
                tx.oncomplete = resolve;
                tx.onerror = () => reject(tx.error);
            });
        } catch (e) {
            console.error('[IndexedDB] Failed to clear data', e);
        }
    }

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
                    if (data.compressedData) {
                        try {
                            const decompressed = LZString.decompressFromUTF16(data.compressedData);
                            window.sharedMissedUnique = JSON.parse(decompressed) || [];
                        } catch (e) {
                            console.error("Failed to decompress shared data:", e);
                            window.sharedMissedUnique = [];
                        }
                    } else {
                        window.sharedMissedUnique = data.missedUnique || [];
                    }
                    // Allow recipient to navigate freely to CO and WOSG sections
                    isAuthenticated = true;
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
        invoice: ['invoice number', 'invoice numb', 'invoice no', 'invoice', 'bill no', 'bill number', 'bill no.', 'invoice no.', 'inv no', 'receipt no'],
        invoiceDate: ['invoice date', 'date', 'bill date', 'creation date'],
        time: ['time', 'creation time', 'invoice time', 'bill time'],
        customerName: ['customer name', 'customer', 'cust name', 'buyer name', 'buyer', 'party name', 'party', 'client name', 'client'],
        customerNo: ['customer number', 'customer no', 'cust no', 'mobile', 'phone', 'contact', 'mobile no', 'phone no', 'contact no', 'customer mobile', 'cust mobile', 'mobile number'],
        pincode: ['pincode', 'pin code', 'zip code', 'zip', 'postal code', 'customer pincode', 'cust pincode'],
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
        invoice: ['invoice no', 'invoice numb', 'invoice number', 'invoice', 'bill no', 'bill number', 'bill no.', 'invoice no.', 'inv no', 'receipt no'],
        staff: ['staff', 'staff name', 'sales person', 'salesman', 'employee name']
    };

    // ---- DOM REFERENCES ----
    const $ = id => {
        const el = document.getElementById(id);
        if (el) return el;
        // Mock element to gracefully handle removed UI sections without throwing runtime errors
        const safeChain = () => safeProxy;
        const safeProxy = new Proxy({}, {
            get: function (target, prop) {
                if (prop === 'then') return undefined;
                if (prop === 'options') return [{ textContent: 'All' }];
                if (prop === 'value' || prop === 'innerHTML' || prop === 'textContent' || prop === 'className') return '';
                if (prop === 'style' || prop === 'classList' || prop === 'dataset') return safeProxy;
                if (typeof prop === 'string') return safeChain;
                return undefined;
            },
            set: function () { return true; }
        });
        return safeProxy;
    };

    const originalQuerySelector = document.querySelector.bind(document);
    document.querySelector = function (selector) {
        const el = originalQuerySelector(selector);
        if (el) return el;
        const safeChain = () => safeProxy;
        const safeProxy = new Proxy({}, {
            get: function (target, prop) {
                if (prop === 'then') return undefined;
                if (prop === 'options') return [{ textContent: 'All' }];
                if (prop === 'value' || prop === 'innerHTML' || prop === 'textContent' || prop === 'className') return '';
                if (prop === 'style' || prop === 'classList' || prop === 'dataset') return safeProxy;
                if (typeof prop === 'string') return safeChain;
                return undefined;
            },
            set: function () { return true; }
        });
        return safeProxy;
    };
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
    const uploadZoneSamsung = $('uploadZoneSamsung');
    const fileInputProduct = $('fileInputProduct');
    const fileInputOSG = $('fileInputOSG');
    const fileInputAMC = $('fileInputAMC');
    const fileInputSamsung = $('fileInputSamsung');
    const productStatus = $('productStatus');
    const osgStatus = $('osgStatus');
    const amcStatus = $('amcStatus');
    const samsungStatus = $('samsungStatus');

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
    let coStatusMap = {};
    function fbKey(inv) {
        return String(inv || '').replace(/[.#$\[\]\/]/g, '_');
    }
    const CO_CALLERS = [
        { name: 'Harmiya', color: '#7c3aed', bg: 'rgba(124,58,237,0.15)', pass: '1234' },
        { name: 'Aswathi', color: '#0891b2', bg: 'rgba(8,145,178,0.15)', pass: '5678' },
        { name: 'Shikha', color: '#d97706', bg: 'rgba(217,119,6,0.15)', pass: '9012' },
        { name: 'Anjana', color: '#e11d48', bg: 'rgba(225,29,72,0.15)', pass: '3456' }
    ];
    let currentCaller = localStorage.getItem('co_caller') || null;
    let coCurrentRows = [];
    let coCurrentSelDate = '';   // current filtered rows
    let coDisplayLimit = 100; // pagination: rows shown so far

    // ---- GLOBAL DOWNLOAD STAFF DETAILS ----
    window.downloadStaffDetails = function (staff, branch) {
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
            const sRow = pRows.find(r => r.staff === s);
            const sBranch = (sRow && sRow.branch) ? sRow.branch : (branch || 'Unknown');

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

    window.downloadBranchDetails = function (branch) {
        window.downloadStaffDetails(null, branch);
    };

    // ---- NAVIGATION ----
    document.querySelectorAll('.nav-item').forEach(item => {
        item.addEventListener('click', (e) => {
            e.preventDefault();
            const section = item.dataset.section;

            // Check if page needs auth ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â Upload Data and Customers Without OSG are public
            const isPublicPage = section === 'customers-osg-section' || section === 'upload-section' || section === 'wosg-dashboard-section';
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

    // ---- TRENDS BUTTON WIRING ----
    // Set default month to current month
    const trendMonthInput = document.getElementById('trendMonthSelect');
    if (trendMonthInput) {
        const now = new Date();
        trendMonthInput.value = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
    }

    document.getElementById('btnSaveTrendMonth').addEventListener('click', async () => {
        const month = document.getElementById('trendMonthSelect').value;
        if (!month) { alert('Please select a month first.'); return; }
        if (!filteredProduct || filteredProduct.length === 0) { alert('No data loaded. Upload files first.'); return; }

        const conv = calcConversion(filteredProduct, filteredOSG);
        const totalQty = filteredProduct.reduce((s, r) => s + (r.qty || 0), 0);
        const totalRev = filteredProduct.reduce((s, r) => s + (r.revenue || r.soldPrice || 0), 0);

        const snapshot = {
            month,
            totalQty,
            totalRev,
            osgQty: conv.oQty,
            osgRev: conv.oSoldPrice,
            qtyConv: parseFloat(conv.qtyConv.toFixed(2)),
            valConv: parseFloat(conv.valueConv.toFixed(2)),
            savedAt: new Date().toISOString()
        };
        await saveMonthlyDataToDB(month, snapshot);
        alert(`ÃƒÂ¢Ã…â€œÃ¢â‚¬Â¦ Data for ${month} saved successfully!`);
        renderTrendsDashboard();
    });

    document.getElementById('btnClearTrends').addEventListener('click', async () => {
        if (!confirm('Are you sure you want to clear ALL historical trend data?')) return;
        await clearAllMonthlyDataFromDB();
        renderTrendsDashboard();
        alert('Historical data cleared.');
    });


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
    // Always use fresh getElementById to avoid safeProxy issues
    function initUploadZones() {
        console.log('[UPLOAD] initUploadZones called');
        const zoneSmart = document.getElementById('uploadZoneSmart');
        const inputSmart = document.getElementById('fileInputSmart');
        console.log('[UPLOAD DEBUG] initUploadZones called. zone:', zoneSmart, 'input:', inputSmart);

        const statusProduct = document.getElementById('productStatus');
        const statusOSG = document.getElementById('osgStatus');
        const statusAMC = document.getElementById('amcStatus');
        const statusSamsung = document.getElementById('samsungStatus');

        if (zoneSmart && inputSmart) {
            const onFilesHandler = async (files) => {
                console.log('[UPLOAD] onFiles called with', files.length, 'files:', files.map(f => f.name));
                for (let file of files) {
                    let fname = file.name.toLowerCase();
                    let fileType = null;

                    if (fname.includes('product') || fname.includes('sales')) {
                        fileType = 'product';
                    } else if (fname.includes('amc') || fname.includes('lgamc')) {
                        fileType = 'amc';
                    } else if (fname.includes('samsung') || fname.includes('sam')) {
                        fileType = 'samsung';
                    } else if (fname.includes('osg') || fname.includes('warranty')) {
                        fileType = 'osg';
                    } else {
                        fileType = await new Promise(resolve => {
                            const old = document.getElementById('_fileTypeModal');
                            if (old) old.remove();
                            const modal = document.createElement('div');
                            modal.id = '_fileTypeModal';
                            modal.style.cssText = 'position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.7);z-index:9999;display:flex;align-items:center;justify-content:center;';
                            modal.innerHTML = '<div style="background:var(--bg-card,#1e2433);border:1px solid var(--border,#374151);border-radius:16px;padding:28px 32px;max-width:420px;width:90%;box-shadow:0 20px 60px rgba(0,0,0,0.5);">'
                                + '<h3 style="margin:0 0 8px;color:var(--text-primary,#f1f5f9);font-size:1.1rem;">Select File Type</h3>'
                                + '<p style="margin:0 0 20px;color:var(--text-muted,#94a3b8);font-size:0.85rem;">Could not auto-detect file type for: <strong style="color:var(--text-secondary,#cbd5e1);">' + file.name + '</strong></p>'
                                + '<div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;">'
                                + '<button id="_ftm_product" style="padding:12px;border-radius:10px;border:1px solid rgba(59,130,246,0.4);background:rgba(59,130,246,0.1);color:#60a5fa;cursor:pointer;font-weight:600;font-family:inherit;font-size:0.88rem;">Product / Sales</button>'
                                + '<button id="_ftm_osg" style="padding:12px;border-radius:10px;border:1px solid rgba(16,185,129,0.4);background:rgba(16,185,129,0.1);color:#10b981;cursor:pointer;font-weight:600;font-family:inherit;font-size:0.88rem;">OSG Warranty</button>'
                                + '<button id="_ftm_amc" style="padding:12px;border-radius:10px;border:1px solid rgba(168,85,247,0.4);background:rgba(168,85,247,0.1);color:#a855f7;cursor:pointer;font-weight:600;font-family:inherit;font-size:0.88rem;">LG-AMC</button>'
                                + '<button id="_ftm_samsung" style="padding:12px;border-radius:10px;border:1px solid rgba(251,191,36,0.4);background:rgba(251,191,36,0.1);color:#fbbf24;cursor:pointer;font-weight:600;font-family:inherit;font-size:0.88rem;">Samsung</button>'
                                + '</div>'
                                + '<button id="_ftm_skip" style="margin-top:14px;width:100%;padding:8px;border-radius:8px;border:1px solid var(--border,#374151);background:transparent;color:var(--text-muted,#94a3b8);cursor:pointer;font-family:inherit;font-size:0.82rem;">Skip this file</button>'
                                + '</div>';
                            document.body.appendChild(modal);
                            const pick = (t) => { modal.remove(); resolve(t); };
                            document.getElementById('_ftm_product').onclick = () => pick('product');
                            document.getElementById('_ftm_osg').onclick = () => pick('osg');
                            document.getElementById('_ftm_amc').onclick = () => pick('amc');
                            document.getElementById('_ftm_samsung').onclick = () => pick('samsung');
                            document.getElementById('_ftm_skip').onclick = () => pick(null);
                        });
                    }

                    if (!fileType) continue;

                    showLoading(true);
                    try {
                        if (fileType === 'product') {
                            const rows = await parseProductFile(file);
                            productData = rows;
                            // Auto-detect active month from uploaded data using robust date parsing
                            window.coActiveMonth = (function () {
                                const counts = {};
                                rows.forEach(r => {
                                    let dt = r.invoiceDate || r.time;
                                    if (!dt) return;
                                    let key = '';
                                    if (dt instanceof Date && !isNaN(dt)) {
                                        key = dt.getFullYear() + '-' + String(dt.getMonth() + 1).padStart(2, '0');
                                    } else if (typeof dt === 'number') {
                                        // Excel serial number (days since 1900-01-01)
                                        const dObj = new Date(Math.round((dt - 25569) * 86400 * 1000));
                                        key = dObj.getUTCFullYear() + '-' + String(dObj.getUTCMonth() + 1).padStart(2, '0');
                                    } else if (typeof dt === 'string') {
                                        const s = dt.trim();
                                        if (s.match(/^\d{4}[\/\-]\d{2}[\/\-]\d{2}/)) {
                                            // ISO format: YYYY-MM-DD or YYYY/MM/DD
                                            key = s.substring(0, 4) + '-' + s.substring(5, 7);
                                        } else {
                                            // DD-MMM-YYYY or DD-MMM-YY (e.g. 01-Apr-2026)
                                            const matchAlpha = s.match(/^(\d{1,2})[\/\-\s]+([A-Za-z]{3,})[\/\-\s]+(\d{2,4})/);
                                            if (matchAlpha) {
                                                const yy = matchAlpha[3].length === 2 ? '20' + matchAlpha[3] : matchAlpha[3];
                                                const mName = matchAlpha[2].toLowerCase();
                                                const mIdx = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'].findIndex(m => mName.startsWith(m)) + 1;
                                                if (mIdx > 0) {
                                                    key = yy + '-' + String(mIdx).padStart(2, '0');
                                                }
                                            } else {
                                                // DD/MM/YYYY or DD-MM-YYYY (Indian format)
                                                const match = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
                                                if (match) {
                                                    const yy = match[3].length === 2 ? '20' + match[3] : match[3];
                                                    key = yy + '-' + match[2].padStart(2, '0');
                                                }
                                            }
                                        }
                                    }
                                    if (key) counts[key] = (counts[key] || 0) + 1;
                                });
                                const best = Object.entries(counts).sort((a, b) => b[1] - a[1])[0];
                                return best ? best[0] : new Date().toISOString().substring(0, 7);
                            })();
                            console.log('[Month] Active month detected:', window.coActiveMonth);
                            isCoStatusLive = false; // Force re-subscribe to new month path
                            if (typeof updateMonthSwitcherUI === 'function') updateMonthSwitcherUI();
                            showFileStatus(statusProduct, file.name, rows.length);
                        } else if (fileType === 'amc') {
                            const rows = await parseLGAMCFile(file);
                            amcData = rows;
                            showFileStatus(statusAMC, file.name, rows.length);
                        } else if (fileType === 'samsung') {
                            const rows = await parseSamsungFile(file);
                            samsungData = rows;
                            showFileStatus(statusSamsung, file.name, rows.length);
                        } else if (fileType === 'osg') {
                            const rows = await parseOSGFile(file);
                            osgData = rows;
                            showFileStatus(statusOSG, file.name, rows.length);
                        }
                    } catch (err) {
                        console.error('Error parsing file:', file.name, err);
                        alert('Error reading file: ' + file.name + '\n' + err.message);
                    } finally {
                        showLoading(false);
                    }
                }
                checkGenerateReady();
            };

            setupUploadZone(zoneSmart, inputSmart, onFilesHandler);
        }
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initUploadZones);
    } else {
        initUploadZones();
    }

    function setupUploadZone(zone, input, onFiles) {
        console.log('[UPLOAD] setupUploadZone registered for', input.id);
        // Only open the file picker on bare zone clicks (not on label/button/input)
        // The label's native `for` attribute already opens the input â€”Â no manual click needed
        zone.addEventListener('click', (e) => {
            if (e.target === input) return;
            if (e.target.tagName === 'LABEL' || e.target.closest('label')) return;
            if (e.target.tagName === 'BUTTON' || e.target.closest('button')) return;
            input.click();
        });
        zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag-over'); });
        zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
        zone.addEventListener('drop', async e => {
            e.preventDefault();
            zone.classList.remove('drag-over');
            if (e.dataTransfer.files.length > 0) {
                try {
                    await onFiles(Array.from(e.dataTransfer.files));
                } catch (err) {
                    console.error(err);
                } finally {
                    input.value = '';
                }
            }
        });
        input.addEventListener('change', async () => {
            console.log('[UPLOAD] input change fired! files:', input.files.length);
            if (input.files.length > 0) {
                try {
                    await onFiles(Array.from(input.files));
                } catch (err) {
                    alert('Error processing files: ' + err.message);
                    console.error('[UPLOAD] onFiles error:', err);
                } finally {
                    input.value = '';
                }
            }
        });
    }

    function showFileStatus(el, name, count) {
        if (!el) return;
        el.className = 'upload-status has-data';
        el.innerHTML = '<span class="status-icon" style="color:#10b981;">&#x2714;</span> ' +
            '<span class="status-text">' + name + '</span> ' +
            '<span class="status-count">' + count + ' rows</span>';
    }

    function checkGenerateReady() {
        const btn = document.getElementById('btnGenerate');
        if (btn) btn.disabled = !(productData.length > 0 && osgData.length > 0);
    }

    (function () {
        var btnGen = document.getElementById('btnGenerate');
        if (!btnGen) return;
        btnGen.addEventListener('click', function () {
            showLoading(true);
            setTimeout(async function () {
                try {
                    // Remove returned products (negative qty) and their matching positive invoice pair
                    const cancelledInv = new Set();
                    const pR = [], nR = [];
                    productData.forEach(r => { if ((r.qty || 0) < 0 || (r.soldPrice || 0) < 0) nR.push(r); else pR.push(r); });
                    nR.forEach(nr => {
                        const mIdx = pR.findIndex(pr => !cancelledInv.has(pr.invoice) &&
                            (pr.customerNo === nr.customerNo || pr.customerName === nr.customerName) &&
                            pr.product === nr.product && Math.abs(pr.soldPrice + nr.soldPrice) < 2);
                        if (mIdx !== -1) {
                            cancelledInv.add(pR[mIdx].invoice);
                            cancelledInv.add(nr.invoice);
                        } else {
                            cancelledInv.add(nr.invoice);
                        }
                    });
                    productData = productData.filter(r => !cancelledInv.has(r.invoice) && (r.qty || 0) > 0);

                    allData = [...productData, ...amcData];

                    if (window.coActiveMonth) {
                        await saveMonthlyDataToDB(window.coActiveMonth, {
                            productData,
                            osgData,
                            amcData,
                            samsungData
                        });
                        console.log('[IndexedDB] Saved data for month:', window.coActiveMonth);
                    }

                    var fcb = document.getElementById('fileCountBadge');
                    var fct = document.getElementById('fileCountText');
                    var bShr = document.getElementById('btnShare');
                    var bRst = document.getElementById('btnReset');
                    if (fcb) fcb.style.display = 'flex';
                    if (fct) fct.textContent = allData.length + ' product Ãƒâ€šÃ‚Â· ' + osgData.length + ' OSG';
                    if (bShr) bShr.style.display = 'flex';
                    if (bRst) bRst.style.display = 'flex';

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



    })();

    // ---- SHARE DASHBOARD LOGIC ----
    (function () {
        var btnShareEl = document.getElementById('btnShare');
        if (!btnShareEl) return;
        btnShareEl.addEventListener('click', function () {
            if (productData.length === 0) return alert('Upload data first via Dashboard.');

            // Find missedUnique for the whole dataset
            const osgInvoices = new Set();
            osgData.forEach(r => { if (r.invoice) osgInvoices.add(r.invoice); });
            amcData.forEach(r => { if (r.invoice) osgInvoices.add(r.invoice); });
            samsungData.forEach(r => { if (r.invoice) osgInvoices.add(r.invoice); });
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

                .map(r => ({
                    invoice: r.invoice || '',
                    customerName: r.customerName || '',
                    customerNo: r.customerNo || '',
                    staff: r.staff || '',
                    branch: r.branch || '',
                    product: r.product || '',
                    brand: r.brand || '',
                    rbm: r.rbm || '',
                    bdm: r.bdm || '',
                    invoiceDate: r.invoiceDate || r.time || '',
                    soldPrice: r.soldPrice || 0,
                    qty: r.qty || 0,
                }));

            showLoading(true);
            try {
                const shareRef = firebase.database().ref('shares').push();
                const compressed = LZString.compressToUTF16(JSON.stringify(payload));
                console.log("Compressed share payload from ~" + JSON.stringify(payload).length + " bytes to " + compressed.length + " bytes");
                shareRef.set({ compressedData: compressed, timestamp: Date.now() })
                    .then(() => {
                        showLoading(false);
                        const base = 'https://officialacc2625-web.github.io/rahul/';
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
    })();

    // ---- PARSING ----
    function parseProductFile(file) {
        return parseExcel(file, PRODUCT_COL_MAP, (row, mapping) => {
            const r = {};
            r.branch = strVal(row, mapping.branch);
            r.rbm = strVal(row, mapping.rbm);
            r.bdm = strVal(row, mapping.bdm);
            r.staff = strVal(row, mapping.staff);
            r.product = strVal(row, mapping.product);
            r.category = normalizeProductCategory(strVal(row, mapping.category));
            r.brand = strVal(row, mapping.brand);
            r.invoice = strVal(row, mapping.invoice);
            r.customerName = strVal(row, mapping.customerName);
            r.invoiceDate = getVal(row, mapping.invoiceDate, '');
            r.time = getVal(row, mapping.time, '');
            r.customerNo = strVal(row, mapping.customerNo);
            r.pincode = strVal(row, mapping.pincode);
            r.soldPrice = num(getVal(row, mapping.soldPrice, 0));
            r.taxableVal = num(getVal(row, mapping.taxableVal, 0));
            r.tax = num(getVal(row, mapping.tax, 0));
            r.qty = num(getVal(row, mapping.qty, 0));
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

    // ---- SAMSUNG PRODUCT NAME NORMALIZER ----
    // Maps raw Samsung Care+ product strings ÃƒÂ¢Ã¢â‚¬Â Ã¢â‚¬â„¢ standard category names
    // e.g. "Samsung Care+ EW WM Auto TopLoad 1Year" ÃƒÂ¢Ã¢â‚¬Â Ã¢â‚¬â„¢ "WASHING MACHINE"
    function mapSamsungProductCategory(rawName) {
        if (!rawName) return rawName;
        const name = rawName.toUpperCase();

        // Order matters: check more-specific tokens first
        // Microwave Oven
        if (/\bMWO\b/.test(name) || name.includes('MICROWAVE') || /\bOVEN\b/.test(name)
            || name.includes('CONVECTION OVEN')) return 'MICROWAVE OVEN';
        // Washing Machine (WM, WM Auto, WM Semi, WM F/Load, WM TopLoad, etc.)
        if (/\bWM\b/.test(name) || name.includes('WASHING MACHINE') || name.includes('WASHER')) return 'WASHING MACHINE';
        // Refrigerator (Ref, Ref DC, Ref FF, Ref SBS, etc.)
        if (/\bREF\b/.test(name) || name.includes('REFRIGERATOR') || name.includes('FRIDGE')) return 'REFRIGERATOR';
        // Air Conditioner
        if (/\bAC\b/.test(name) || name.includes('AIR CONDITIONER') || name.includes('AIRCONDITIONER')) return 'AC';
        // TV / Display
        if (/\bTV\b/.test(name) || name.includes('TELEVISION') || name.includes('MONITOR')) return 'TV';

        // If no abbreviation matched, return as-is (uppercase for consistency)
        return rawName;
    }

    // ---- UNIVERSAL CATEGORY NORMALIZER (for Product & OSG files) ----
    // Maps raw category column values ÃƒÂ¢Ã¢â‚¬Â Ã¢â‚¬â„¢ one of the 9 standard categories
    function normalizeProductCategory(raw) {
        if (!raw) return raw;
        const c = raw.toUpperCase().trim();
        if (c.includes('MICROWAVE') || c.includes('MWO') || c.includes('OVEN')) return 'MICROWAVE OVEN';
        if (c.includes('WASHING MACHINE') || c.includes('WASHER') || /\bWM\b/.test(c)) return 'WASHING MACHINE';
        if (c.includes('DRYER')) return 'DRYER';
        if (c.includes('REFRIGERATOR') || c.includes('FRIDGE') || c.includes('FROST FREE')
            || c.includes('DIRECT COOL') || c.includes('SIDE BY SIDE') || /\bREF\b/.test(c)) return 'REFRIGERATOR';
        if (c.includes('AIR CONDITIONER') || c.includes('SPLIT') || c.includes('WINDOW AC')
            || /\bAC\b/.test(c)) return 'AC';
        if (c.includes('TELEVISION') || c.includes('TV') || c.includes('LED') || c.includes('OLED')
            || c.includes('MONITOR') || c.includes('DISPLAY')) return 'TV';
        if (c.includes('AUDIO') || c.includes('SOUND') || c.includes('SPEAKER')
            || c.includes('HOME THEATER') || c.includes('HOME THEATRE')) return 'AUDIO SYSTEM';
        if (c.includes('CHIMNEY') || c.includes('HOB') || c.includes('INDUCTION')
            || c.includes('VACUUM') || c.includes('VACCUM') || c.includes('WATER PURIFIER')
            || c.includes('PURIFIER') || c.includes('WATER HEATER') || c.includes('GEYSER')
            || c.includes('HOME APPLIANCE')) return 'HOME APPLIANCE';
        if (c.includes('DISH WASHER') || c.includes('DISHWASHER')) return 'DISH WASHER';
        // Return the original if not matched â€”Â keeps custom categories intact
        return raw;
    }

    // ---- LG-AMC PRODUCT NAME NORMALIZER ----
    // Maps full LG product names ÃƒÂ¢Ã¢â‚¬Â Ã¢â‚¬â„¢ standard category names.
    // Audio-related ÃƒÂ¢Ã¢â‚¬Â Ã¢â‚¬â„¢ AUDIO SYSTEM; anything else not in the 6 known categories ÃƒÂ¢Ã¢â‚¬Â Ã¢â‚¬â„¢ SMALL APPLIANCE
    function mapLGAMCProductCategory(rawName) {
        if (!rawName) return rawName;
        const name = rawName.toUpperCase();

        // --- Known major 6 categories first ---
        // Microwave Oven (includes LG Oven Convection, etc.)
        if (/\bMWO\b/.test(name) || name.includes('MICROWAVE') || /\bOVEN\b/.test(name)
            || name.includes('CONVECTION')) return 'MICROWAVE OVEN';
        // Washing Machine
        if (/\bWM\b/.test(name) || name.includes('WASHING MACHINE') || name.includes('WASHER') || name.includes('WASH')
            || name.includes('FRONT LOAD') || name.includes('TOP LOAD') || name.includes('FL WM')
            || name.includes('TL WM') || name.includes('F/L') || name.includes('T/L') 
            || /\bSA\b/.test(name) || name.includes('SEMI') || name.includes('FL DRYER') || name.includes('DRYER')) {
            // Dryer is laundry but separate â€”Â return DRYER
            if (name.includes('DRYER') && !name.includes('WASHER') && !name.includes('WASHING')) return 'DRYER';
            return 'WASHING MACHINE';
        }
        // Refrigerator
        if (/\bREF\b/.test(name) || name.includes('REFRIGERATOR') || name.includes('FRIDGE')
            || name.includes('SIDE BY SIDE') || /\bSBS\b/.test(name) || /\bFFR\b/.test(name) || /\bFF\b/.test(name)
            || name.includes('FROST FREE') || name.includes('DIRECT COOL') || /\bDC\b/.test(name)) return 'REFRIGERATOR';
        // Air Conditioner
        if (/\bAC\b/.test(name) || name.includes('AIR CONDITIONER') || name.includes('SPLIT')
            || name.includes('WINDOW') || name.includes('INVERTER AC')) return 'AC';
        // TV / Display
        if (/\bTV\b/.test(name) || name.includes('TELEVISION') || name.includes('OLED')
            || name.includes('QNED') || name.includes('NANOCELL') || name.includes('SMART TV')
            || /\bLED\b/.test(name) || /\bUHD\b/.test(name) || /\bFHD\b/.test(name) || /\bHD\b/.test(name)
            || name.includes('MONITOR')) return 'TV';

        // --- Audio System ---
        if (name.includes('SOUND BAR') || name.includes('SOUNDBAR') || name.includes('XBOOM')
            || name.includes('SPEAKER') || name.includes('HOME THEATER') || name.includes('HOME THEATRE')
            || name.includes('AUDIO') || name.includes('SUBWOOFER') || name.includes('WOOFER')
            || name.includes('HI-FI') || name.includes('HIFI') || name.includes('STEREO')) return 'AUDIO SYSTEM';

        // --- Home Appliance (Chimney, Hob, Vacuum, Water Purifier/Heater) ---
        if (/\bWP\b/.test(name) || name.includes('WATER PURIFIER') || name.includes('PURIFIER')
            || /\bRO\b/.test(name) || name.includes('WATER PURIF')
            || name.includes('VACUUM') || name.includes('VACCUM') || name.includes('VAC CLEANER')
            || name.includes('CORDZERO') || name.includes('CORDLESS STICK') || name.includes('ROBO')
            || name.includes('CHIMNEY') || /\bHOB\b/.test(name) || name.includes('INDUCTION')
            || name.includes('WATER HEATER') || name.includes('GEYSER')) {
            return 'HOME APPLIANCE';
        }

        // --- Dish Washer ---
        if (/\bDW\b/.test(name) || name.includes('DISHWASHER') || name.includes('DISH WASHER')) return 'DISH WASHER';

        // --- Air Purifier ---
        if (name.includes('AIR PURIFIER') || name.includes('AIR PURIF') || /\bAP\b/.test(name)) return 'AIR PURIFIER';

        // --- Laptop ---
        if (name.includes('LAPTOP') || name.includes('GRAM') || name.includes('NOTEBOOK')) return 'LAPTOP';

        // --- Everything else ÃƒÂ¢Ã¢â‚¬Â Ã¢â‚¬â„¢ Small Appliance ---
        // (Styler, WashTower, other niche products)
        return 'SMALL APPLIANCE';
    }

    function parseOSGFile(file) {
        return parseExcel(file, OSG_COL_MAP, (row, mapping) => {
            const r = {};
            r.branch = strVal(row, mapping.branch);
            r.storeCode = strVal(row, mapping.storeCode);
            r.product = strVal(row, mapping.product);
            r.category = normalizeProductCategory(strVal(row, mapping.category));
            r.brand = strVal(row, mapping.brand);
            r.soldPrice = num(getVal(row, mapping.soldPrice, 0));
            r.qty = parseFloat(strVal(row, mapping.qty)) || 1;
            r.invoice = strVal(row, mapping.invoice);
            r.staff = strVal(row, mapping.staff);

            // If no branch column was found in the file, the "Brand" column likely
            // holds the branch/store name (common in OSG exports with duplicate "Brand" headers).
            if (!mapping.branch && !r.branch && r.brand) {
                r.branch = r.brand;
                r.brand = ''; // Clear since this wasn't actually the brand
            }

            // Ensure category is always normalized from the product column
            if (!r.category || r.category === 'SMALL APPLIANCE') {
                const normalized = normalizeProductCategory(r.product);
                if (normalized !== 'SMALL APPLIANCE' || !r.category) {
                    r.category = normalized;
                }
            }

            return r;
        }).then(rows => {
            // DEBUG: Log OSG parsing results
            console.log('=== OSG DEBUG ===');
            console.log('Total OSG rows parsed:', rows.length);
            const withBranch = rows.filter(r => r.branch).length;
            const withInvoice = rows.filter(r => r.invoice).length;
            const withProduct = rows.filter(r => r.product).length;
            const futureBranches = rows.filter(r => r.branch && r.branch.toUpperCase().includes('FUTURE')).length;
            console.log('Rows with branch:', withBranch);
            console.log('Rows with invoice:', withInvoice);
            console.log('Rows with product:', withProduct);
            console.log('Rows with FUTURE in branch:', futureBranches);
            // Show unique branches
            const branches = {};
            rows.forEach(r => { branches[r.branch || '(empty)'] = (branches[r.branch || '(empty)'] || 0) + 1; });
            console.log('OSG branches:', JSON.stringify(branches));
            // Show first 5 rows
            console.log('OSG sample rows:', JSON.stringify(rows.slice(0, 5).map(r => ({branch: r.branch, product: r.product, category: r.category, invoice: r.invoice, brand: r.brand, qty: r.qty}))));
            console.log('=== END OSG DEBUG ===');
            return rows;
        });
    }

    // LG-AMC-specific parser: same as parseProductFile but auto-normalizes product names
    function parseLGAMCFile(file) {
        return parseExcel(file, PRODUCT_COL_MAP, (row, mapping) => {
            const r = {};
            r.branch = strVal(row, mapping.branch);
            r.rbm = strVal(row, mapping.rbm);
            r.bdm = strVal(row, mapping.bdm);
            r.staff = strVal(row, mapping.staff);
            // Normalize the raw LG product name ÃƒÂ¢Ã¢â‚¬Â Ã¢â‚¬â„¢ standard category
            const rawProduct = strVal(row, mapping.product);
            r.product = rawProduct;
            r.rawProduct = rawProduct;  // keep original for debugging
            r.category = mapLGAMCProductCategory(rawProduct);
            r.brand = strVal(row, mapping.brand);
            r.invoice = strVal(row, mapping.invoice);
            r.customerName = strVal(row, mapping.customerName);
            r.invoiceDate = getVal(row, mapping.invoiceDate, '');
            r.time = getVal(row, mapping.time, '');
            r.customerNo = strVal(row, mapping.customerNo);
            r.soldPrice = num(getVal(row, mapping.soldPrice, 0));
            r.taxableVal = num(getVal(row, mapping.taxableVal, 0));
            r.tax = num(getVal(row, mapping.tax, 0));
            r.qty = num(getVal(row, mapping.qty, 0));
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

    // Samsung-specific parser: same as OSG but auto-normalizes product names
    function parseSamsungFile(file) {
        return parseExcel(file, OSG_COL_MAP, (row, mapping) => {
            const r = {};
            r.branch = strVal(row, mapping.branch);
            r.storeCode = strVal(row, mapping.storeCode);
            // Normalize the raw product name ÃƒÂ¢Ã¢â‚¬Â Ã¢â‚¬â„¢ standard category
            const rawProduct = strVal(row, mapping.product);
            r.product = rawProduct;
            r.rawProduct = rawProduct;  // keep original for debugging
            r.category = mapSamsungProductCategory(rawProduct);
            r.brand = strVal(row, mapping.brand);
            r.soldPrice = num(getVal(row, mapping.soldPrice, 0));
            r.qty = num(getVal(row, mapping.qty, 0));
            r.invoice = strVal(row, mapping.invoice);
            r.staff = strVal(row, mapping.staff);

            // If no branch column was found, the "Brand" column likely holds the branch name
            if (!mapping.branch && !r.branch && r.brand) {
                r.branch = r.brand;
                r.brand = '';
            }

            return r;
        });
    }

    function parseExcel(file, colMap, rowMapper) {
        // Detect CSV files and use fast CSV parser
        const isCSV = file.name.toLowerCase().endsWith('.csv');
        if (isCSV) return parseCSVFile(file, colMap, rowMapper);

        // Guard: XLSX library must be loaded from CDN
        if (typeof XLSX === 'undefined') {
            return Promise.reject(new Error(
                'The XLSX library failed to load.\n\n' +
                'ÃƒÂ¢Ã…â€œÃ¢â‚¬Â¦ Fix: Check your internet connection and reload the page.\n' +
                '   The library loads from cdn.jsdelivr.net on first use.\n\n' +
                'ÃƒÂ°Ã…Â¸Ã¢â‚¬â„¢Ã‚Â¡ Alternative: Convert your file to CSV first:\n' +
                '   python convert.py "your_file.xlsx"'
            ));
        }

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
            if (!mapping[k]) console.warn(`[ÃƒÂ¢Ã…Â¡Ã‚Â  Column NOT FOUND] '${k}' ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â no matching header. Available headers:`, headers.join(', '));
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
        // Product dropdown is predefined ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â keep as-is, just reset selection
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

        // Build metadata helpers for fallbacks
        window.branchToMeta = window.branchToMeta || {};
        window.productInvMeta = window.productInvMeta || {};
        productData.forEach(r => {
            if (r.branch && r.rbm && r.bdm) {
                const bUpper = r.branch.toUpperCase().trim();
                if (!window.branchToMeta[bUpper]) {
                    window.branchToMeta[bUpper] = { rbm: r.rbm, bdm: r.bdm, origBranch: r.branch };
                }
            }
            if (r.invoice) {
                window.productInvMeta[r.invoice] = r;
            }
        });

        // Helper to robustly find a branch meta by exact or fuzzy match, with invoice fallback
        window.getBranchMeta = function(r) {
            let b = null;
            if (r && r.invoice && window.productInvMeta && window.productInvMeta[r.invoice]) {
                const pRow = window.productInvMeta[r.invoice];
                if (pRow.branch) b = pRow.branch.toUpperCase().trim();
            }
            if (!b && r && r.branch) {
                b = r.branch.toUpperCase().trim();
            }
            
            if (!b || !window.branchToMeta) return null;
            
            if (window.branchToMeta[b]) return window.branchToMeta[b];
            
            // Fuzzy match (e.g. 'FALNIR' matches 'FALNIR FUTURE')
            for (const key of Object.keys(window.branchToMeta)) {
                if (key.includes(b) || b.includes(key)) {
                    return window.branchToMeta[key];
                }
            }
            return null;
        };


        // Filter OSG data by product, brand, branch (OSG has these fields)
        const hasPersonFilter = fRBM || fBDM || fStaff;
        const productInvoices = new Set();
        if (hasPersonFilter) {
            filteredProduct.forEach(r => { if (r.invoice) productInvoices.add(r.invoice); });

            filteredOSG = osgData.filter(r => {
                if (fBranch && r.branch !== fBranch) return false;
                if (fProduct && r.product !== fProduct) return false;
                if (fBrand && r.brand !== fBrand) return false;
                
                let rbm = r.rbm || '';
                let bdm = r.bdm || '';
                let staff = r.staff || '';
                
                if (r.invoice && window.productInvMeta[r.invoice]) {
                    rbm = window.productInvMeta[r.invoice].rbm || rbm;
                    bdm = window.productInvMeta[r.invoice].bdm || bdm;
                    staff = window.productInvMeta[r.invoice].staff || staff;
                } else if (window.getBranchMeta(r)) {
                    const m = window.getBranchMeta(r);
                    rbm = rbm || m.rbm;
                    bdm = bdm || m.bdm;
                }
                
                if (fRBM && rbm !== fRBM) return false;
                if (fBDM && bdm !== fBDM) return false;
                if (fStaff && staff !== fStaff) return false;
                
                if (!rbm && !bdm && !staff && !r.invoice) return true;
                if (!rbm && !bdm && !staff) return false; // unmatched completely
                return true;
            });
        } else {
            filteredOSG = osgData.filter(r => {
                if (fBranch && r.branch !== fBranch) return false;
                if (fProduct && r.product !== fProduct) return false;
                if (fBrand && r.brand !== fBrand) return false;
                return true;
            });
        }

        // Filter AMC data: respects branch, product/category, brand dropdowns
        filteredAMC = [];
        if (amcData && amcData.length > 0) {
            filteredAMC = amcData.filter(r => {
                if (fBranch && r.branch !== fBranch) return false;
                if (fProduct && r.product !== fProduct && r.category !== fProduct) return false;
                if (fBrand && r.brand !== fBrand) return false;
                if (hasPersonFilter) {
                    // Try to match using AMC row's own rbm/bdm/staff attributes first
                    let rbm = r.rbm || '';
                    let bdm = r.bdm || '';
                    let staff = r.staff || '';
                    
                    // If the row lacks direct staff info, fallback to productInvoices lookup
                    if (!rbm && !bdm && !staff) {
                        if (r.invoice && productInvoices.has(r.invoice)) return true;
                        if (!r.invoice) return true; // keep care plans without invoices just in case
                        return false;
                    }
                    
                    // Direct match checks
                    if (fRBM && rbm !== fRBM) return false;
                    if (fBDM && bdm !== fBDM) return false;
                    if (fStaff && staff !== fStaff) return false;
                }
                return true;
            });
        }

        // Filter Samsung data: respects branch, product/category, brand dropdowns
        filteredSamsung = [];
        if (samsungData && samsungData.length > 0) {
            filteredSamsung = samsungData.filter(r => {
                if (fBranch && r.branch !== fBranch) return false;
                if (fProduct && r.product !== fProduct && r.category !== fProduct) return false;
                if (fBrand && r.brand !== fBrand) return false;
                if (hasPersonFilter) {
                    let rbm = r.rbm || '';
                    let bdm = r.bdm || '';
                    let staff = r.staff || '';
                    
                    if (r.invoice && window.productInvMeta[r.invoice]) {
                        rbm = window.productInvMeta[r.invoice].rbm || rbm;
                        bdm = window.productInvMeta[r.invoice].bdm || bdm;
                        staff = window.productInvMeta[r.invoice].staff || staff;
                    } else if (window.getBranchMeta(r)) {
                        const m = window.getBranchMeta(r);
                        rbm = rbm || m.rbm;
                        bdm = bdm || m.bdm;
                    }
                    
                    if (fRBM && rbm !== fRBM) return false;
                    if (fBDM && bdm !== fBDM) return false;
                    if (fStaff && staff !== fStaff) return false;
                    
                    if (!rbm && !bdm && !staff && !r.invoice) return true;
                    if (!rbm && !bdm && !staff) return false;
                }
                return true;
            });
        }

        renderDashboard();
        renderReports();
        renderCharts();

        // Show PDF button when data is loaded
        const pdfBtn = document.getElementById('btnPdfReport');
        if (pdfBtn) pdfBtn.style.display = filteredProduct.length > 0 ? 'inline-flex' : 'none';
    }

    // ---- PDF EXECUTIVE SUMMARY ----
    document.getElementById('btnPdfReport').addEventListener('click', generatePDFReport);

    function generatePDFReport() {
        if (!filteredProduct || filteredProduct.length === 0) { alert('No data loaded.'); return; }

        const conv = calcConversion(filteredProduct, filteredOSG);
        const totalQty = filteredProduct.reduce((s, r) => s + (r.qty || 0), 0);
        const totalRev = filteredProduct.reduce((s, r) => s + (r.revenue || r.soldPrice || 0), 0);
        const osgQty = conv.oQty;
        const osgRev = conv.oSoldPrice;

        // LG AMC stats
        const lgQty = filteredProduct.reduce((s, r) => s + ((r.brand && r.brand.toUpperCase().includes('LG')) ? (r.qty || 0) : 0), 0);
        const amcQty = filteredAMC.reduce((s, r) => s + (r.qty || 0), 0);
        const amcConv = lgQty > 0 ? ((amcQty / lgQty) * 100).toFixed(2) : '0.00';

        // Samsung stats
        const samCats = ['AC','MICROWAVE OVEN','REFRIGERATOR','WASHING MACHINE','TV'];
        const samProdQty = filteredProduct.reduce((s, r) => { const p = (r.product || '').toUpperCase().trim(); return s + ((r.brand && r.brand.toUpperCase().includes('SAMSUNG') && samCats.includes(p)) ? (r.qty || 0) : 0); }, 0);
        const samOsgQty = filteredSamsung.reduce((s, r) => s + (r.qty || 0), 0);
        const samConv = samProdQty > 0 ? ((samOsgQty / samProdQty) * 100).toFixed(2) : '0.00';

        // Top 5 branches by revenue
        const branchMap = {};
        filteredProduct.forEach(r => {
            const b = r.branch || 'Unknown';
            if (!branchMap[b]) branchMap[b] = { qty: 0, rev: 0 };
            branchMap[b].qty += (r.qty || 0);
            branchMap[b].rev += (r.revenue || r.soldPrice || 0);
        });
        const topBranches = Object.entries(branchMap).sort((a, b) => b[1].rev - a[1].rev).slice(0, 5);

        // Top 5 staff by qty
        const staffMap = {};
        filteredProduct.forEach(r => {
            const st = r.staff || 'Unknown';
            if (!staffMap[st]) staffMap[st] = { qty: 0 };
            staffMap[st].qty += (r.qty || 0);
        });
        const topStaff = Object.entries(staffMap).sort((a, b) => b[1].qty - a[1].qty).slice(0, 5);

        const today = new Date().toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' });

        const reportHTML = `
        <div style="font-family:'Segoe UI',system-ui,sans-serif; color:#1e293b; padding:30px; max-width:800px; margin:auto;">
            <div style="text-align:center; margin-bottom:30px; border-bottom:3px solid #f97316; padding-bottom:20px;">
                <h1 style="margin:0; font-size:24px; color:#f97316; letter-spacing:-0.5px;">CRM-DATA ANALYTICS PORTAL</h1>
                <h2 style="margin:6px 0 0; font-size:16px; color:#64748b; font-weight:400;">Executive Summary Report &bull; ${today}</h2>
            </div>

            <h3 style="color:#f97316; border-bottom:2px solid #fed7aa; padding-bottom:6px; margin-bottom:12px;">Key Performance Indicators</h3>
            <table style="width:100%; border-collapse:collapse; margin-bottom:24px; font-size:13px;">
                <tr style="background:#f8fafc;">
                    <td style="padding:10px 12px; border:1px solid #e2e8f0; font-weight:600;">Total Product Qty</td>
                    <td style="padding:10px 12px; border:1px solid #e2e8f0; text-align:right;">${formatNumber(totalQty)}</td>
                    <td style="padding:10px 12px; border:1px solid #e2e8f0; font-weight:600;">Total Product Revenue</td>
                    <td style="padding:10px 12px; border:1px solid #e2e8f0; text-align:right;">\u20B9${fmtShortHtml(totalRev)}</td>
                </tr>
                <tr>
                    <td style="padding:10px 12px; border:1px solid #e2e8f0; font-weight:600;">OSG Qty</td>
                    <td style="padding:10px 12px; border:1px solid #e2e8f0; text-align:right;">${formatNumber(osgQty)}</td>
                    <td style="padding:10px 12px; border:1px solid #e2e8f0; font-weight:600;">OSG Revenue</td>
                    <td style="padding:10px 12px; border:1px solid #e2e8f0; text-align:right;">\u20B9${fmtShortHtml(osgRev)}</td>
                </tr>
                <tr style="background:#f8fafc;">
                    <td style="padding:10px 12px; border:1px solid #e2e8f0; font-weight:600;">Qty Conversion</td>
                    <td style="padding:10px 12px; border:1px solid #e2e8f0; text-align:right; color:${conv.qtyConv >= 15 ? '#10b981' : '#ef4444'}; font-weight:700;">${conv.qtyConv.toFixed(2)}%</td>
                    <td style="padding:10px 12px; border:1px solid #e2e8f0; font-weight:600;">Value Conversion</td>
                    <td style="padding:10px 12px; border:1px solid #e2e8f0; text-align:right; color:${conv.valueConv >= 15 ? '#10b981' : '#ef4444'}; font-weight:700;">${conv.valueConv.toFixed(2)}%</td>
                </tr>
                <tr>
                    <td style="padding:10px 12px; border:1px solid #e2e8f0; font-weight:600;">LG-AMC Conv%</td>
                    <td style="padding:10px 12px; border:1px solid #e2e8f0; text-align:right; font-weight:700;">${amcConv}%</td>
                    <td style="padding:10px 12px; border:1px solid #e2e8f0; font-weight:600;">Samsung Conv%</td>
                    <td style="padding:10px 12px; border:1px solid #e2e8f0; text-align:right; font-weight:700;">${samConv}%</td>
                </tr>
            </table>

            <h3 style="color:#3b82f6; border-bottom:2px solid #bfdbfe; padding-bottom:6px; margin-bottom:12px;">Top 5 Branches by Revenue</h3>
            <table style="width:100%; border-collapse:collapse; margin-bottom:24px; font-size:13px;">
                <tr style="background:#1e293b; color:#fff;">
                    <th style="padding:8px 12px; text-align:left;">Branch</th>
                    <th style="padding:8px 12px; text-align:right;">Qty</th>
                    <th style="padding:8px 12px; text-align:right;">Revenue</th>
                </tr>
                ${topBranches.map(([name, d], i) => `
                    <tr style="background:${i % 2 === 0 ? '#f8fafc' : '#fff'};">
                        <td style="padding:8px 12px; border:1px solid #e2e8f0; font-weight:600;">${name}</td>
                        <td style="padding:8px 12px; border:1px solid #e2e8f0; text-align:right;">${formatNumber(d.qty)}</td>
                        <td style="padding:8px 12px; border:1px solid #e2e8f0; text-align:right;">\u20B9${fmtShortHtml(d.rev)}</td>
                    </tr>
                `).join('')}
            </table>

            <h3 style="color:#10b981; border-bottom:2px solid #a7f3d0; padding-bottom:6px; margin-bottom:12px;">Top 5 Staff by Volume</h3>
            <table style="width:100%; border-collapse:collapse; margin-bottom:24px; font-size:13px;">
                <tr style="background:#1e293b; color:#fff;">
                    <th style="padding:8px 12px; text-align:left;">Staff</th>
                    <th style="padding:8px 12px; text-align:right;">Qty Sold</th>
                </tr>
                ${topStaff.map(([name, d], i) => `
                    <tr style="background:${i % 2 === 0 ? '#f8fafc' : '#fff'};">
                        <td style="padding:8px 12px; border:1px solid #e2e8f0; font-weight:600;">${name}</td>
                        <td style="padding:8px 12px; border:1px solid #e2e8f0; text-align:right;">${formatNumber(d.qty)}</td>
                    </tr>
                `).join('')}
            </table>

            <div style="text-align:center; color:#94a3b8; font-size:11px; margin-top:30px; border-top:1px solid #e2e8f0; padding-top:12px;">
                Generated by CRM-DATA ANALYTICS PORTAL &bull; ${today} &bull; Confidential
            </div>
        </div>`;

        const container = document.createElement('div');
        container.innerHTML = reportHTML;

        const opt = {
            margin: [10, 10, 10, 10],
            filename: `Executive_Summary_${today.replace(/\s/g, '_')}.pdf`,
            image: { type: 'jpeg', quality: 0.98 },
            html2canvas: { scale: 2, useCORS: true },
            jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
        };

        html2pdf().set(opt).from(container).save();
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
        const data = filteredProduct;
        const totalQty = data.reduce((s, r) => s + r.qty, 0);

        const conv = calcConversion(filteredProduct, filteredOSG);

        $('kpiValConv').textContent = conv.valueConv.toFixed(2) + '%';
        $('kpiQtyConv').textContent = conv.qtyConv.toFixed(2) + '%';
        $('kpiQuantity').textContent = formatNumber(totalQty);

        // Total Product Sale (revenue)
        $('kpiProdRevenue').textContent = fmtShort(conv.pSoldPrice);
        // Total OSG Sale (revenue)
        $('kpiOsgRevenue').textContent = fmtShort(conv.oSoldPrice);
        // Total OSG Qty
        $('kpiOsgQty').textContent = formatNumber(conv.oQty);
        // Without OSG: total product qty on invoices that have no match in OSG data
        const osgInvSet = new Set();
        filteredOSG.forEach(r => { if (r.invoice) osgInvSet.add(r.invoice); });
        let withoutOsgQty = 0;
        filteredProduct.forEach(r => {
            if (r.invoice && !osgInvSet.has(r.invoice)) {
                withoutOsgQty += (r.qty || 1);
            }
        });
        $('kpiWithoutOsg').textContent = formatNumber(withoutOsgQty);

        // Conversion breakdown tables
        renderConvTable('convRBMTable', 'rbm');
        renderConvTable('convBDMTable', 'bdm');
        renderConvTable('convStaffTable', 'staff');
        renderConvTable('convBranchTable', 'branch');
        renderConvTable('convProductTable', 'product');

        // ---- LG AMC KPI section (show only when amcData is loaded) ----
        const lgAmcKpiRow = document.getElementById('lgAmcKpiRow');
        if (amcData && amcData.length > 0 && lgAmcKpiRow) {
            lgAmcKpiRow.style.display = 'block';

            const lgTotalQty = filteredProduct.reduce((s, r) => s + ((r.brand && r.brand.toUpperCase().includes('LG')) ? (r.qty || 0) : 0), 0);
            const amcTotalQty = filteredAMC.reduce((s, r) => s + (r.qty || 0), 0);
            const amcTotalSale = filteredAMC.reduce((s, r) => s + (r.soldPrice || 0), 0);
            const withoutAmcQty = Math.max(0, lgTotalQty - amcTotalQty);
            const amcConvPct = lgTotalQty > 0 ? (amcTotalQty / lgTotalQty) * 100 : 0;

            const kpiAmcTotal = document.getElementById('kpiAmcTotal');
            const kpiAmcSale = document.getElementById('kpiAmcSale');
            const kpiAmcWithout = document.getElementById('kpiAmcWithout');
            const kpiAmcConv = document.getElementById('kpiAmcConv');

            if (kpiAmcTotal) kpiAmcTotal.textContent = formatNumber(amcTotalQty);
            if (kpiAmcSale) kpiAmcSale.textContent = '\u20B9' + fmtShort(amcTotalSale);
            if (kpiAmcWithout) kpiAmcWithout.textContent = formatNumber(withoutAmcQty);
            if (kpiAmcConv) kpiAmcConv.textContent = amcConvPct.toFixed(2) + '%';
        } else if (lgAmcKpiRow) {
            lgAmcKpiRow.style.display = 'none';
        }

        // ---- Samsung OSG KPI section (show only when samsungData is loaded) ----
        const samsungKpiRow = document.getElementById('samsungKpiRow');
        if (samsungData && samsungData.length > 0 && samsungKpiRow) {
            samsungKpiRow.style.display = 'block';

            const samsungAllowedCats = ['AC', 'MICROWAVE OVEN', 'REFRIGERATOR', 'WASHING MACHINE', 'TV'];

            const samsungBrandTotalQty = filteredProduct.reduce((s, r) => {
                const p = r.product ? r.product.toUpperCase().trim() : '';
                return s + ((r.brand && r.brand.toUpperCase().includes('SAMSUNG') && samsungAllowedCats.includes(p)) ? (r.qty || 0) : 0);
            }, 0);

            const samsungOsgTotalQty = filteredSamsung.reduce((s, r) => {
                const p = r.product ? r.product.toUpperCase().trim() : (r.category ? r.category.toUpperCase().trim() : '');
                return s + ((samsungAllowedCats.includes(p)) ? (r.qty || 0) : 0);
            }, 0);
            const samsungOsgTotalSale = filteredSamsung.reduce((s, r) => {
                const p = r.product ? r.product.toUpperCase().trim() : (r.category ? r.category.toUpperCase().trim() : '');
                return s + ((samsungAllowedCats.includes(p)) ? (r.soldPrice || 0) : 0);
            }, 0);
            const withoutSamsungQty = Math.max(0, samsungBrandTotalQty - samsungOsgTotalQty);
            const samsungConvPct = samsungBrandTotalQty > 0 ? (samsungOsgTotalQty / samsungBrandTotalQty) * 100 : 0;

            const kpiSamsungTotal = document.getElementById('kpiSamsungTotal');
            const kpiSamsungSale = document.getElementById('kpiSamsungSale');
            const kpiSamsungWithout = document.getElementById('kpiSamsungWithout');
            const kpiSamsungConv = document.getElementById('kpiSamsungConv');

            if (kpiSamsungTotal) kpiSamsungTotal.textContent = formatNumber(samsungOsgTotalQty);
            if (kpiSamsungSale) kpiSamsungSale.textContent = '\u20B9' + fmtShort(samsungOsgTotalSale);
            if (kpiSamsungWithout) kpiSamsungWithout.textContent = formatNumber(withoutSamsungQty);
            if (kpiSamsungConv) kpiSamsungConv.textContent = samsungConvPct.toFixed(2) + '%';
        } else if (samsungKpiRow) {
            samsungKpiRow.style.display = 'none';
        }

        renderPincodeDashboard();
        renderProductDashboard();
    }

    // ---- PRODUCT ANALYTICS CHART INSTANCES ----
    const prodChartInstances = {};

    function renderProductDashboard() {
        if (!filteredProduct || filteredProduct.length === 0) {
            const noData = document.getElementById('prodNoData');
            const tbody = document.getElementById('prodDetailTbody');
            if (noData) noData.style.display = 'flex';
            if (tbody) tbody.innerHTML = '';
            return;
        }

        // ----- Aggregate by category -----
        const invMapOSG = {};
        filteredProduct.forEach(r => { if (r.invoice) invMapOSG[r.invoice] = r.product || 'Unknown'; });

        const catMap = {}; // key: category
        filteredProduct.forEach(r => {
            const cat = (r.product || r.category || 'Unknown').toUpperCase().trim() || 'Unknown';
            if (!catMap[cat]) catMap[cat] = { pQty: 0, pRev: 0, oQty: 0, aQty: 0, lgPQty: 0 };
            catMap[cat].pQty += (r.qty || 0);
            catMap[cat].pRev += (r.revenue || r.soldPrice || 0);
            if (r.brand && r.brand.toUpperCase().includes('LG')) catMap[cat].lgPQty += (r.qty || 0);
        });

        filteredOSG.forEach(r => {
            const cat = (r.product || r.category || 'Unknown').toUpperCase().trim();
            if (catMap[cat]) catMap[cat].oQty += (r.qty || 0);
            else {
                // try matching via invoice
                const invoice = r.invoice;
                const pCat = invoice && invMapOSG[invoice] ? (invMapOSG[invoice]).toUpperCase().trim() : null;
                if (pCat && catMap[pCat]) catMap[pCat].oQty += (r.qty || 0);
            }
        });

        filteredAMC.forEach(r => {
            const cat = (r.product || 'Unknown').toUpperCase().trim();
            const branch = r.branch || 'Unknown';
            const key = cat + '|' + branch;
            if (catMap[cat]) catMap[cat].aQty += (r.qty || 0);
        });

        const cats = Object.entries(catMap)
            .map(([name, d]) => ({
                name,
                pQty: d.pQty, pRev: d.pRev, oQty: d.oQty, aQty: d.aQty, lgPQty: d.lgPQty,
                osgConv: d.pQty > 0 ? (d.oQty / d.pQty) * 100 : 0,
                amcConv: d.lgPQty > 0 ? (d.aQty / d.lgPQty) * 100 : 0,
                avgPrice: d.pQty > 0 ? d.pRev / d.pQty : 0
            }))
            .sort((a, b) => b.pRev - a.pRev);

        // ----- KPI Cards -----
        const totalQty = cats.reduce((s, c) => s + c.pQty, 0);
        const totalRev = cats.reduce((s, c) => s + c.pRev, 0);
        const avgConv = cats.length > 0 ? cats.reduce((s, c) => s + c.osgConv, 0) / cats.length : 0;

        const elQ = document.getElementById('prodKpiTotalQty');
        const elR = document.getElementById('prodKpiTotalRev');
        const elC = document.getElementById('prodKpiCategories');
        const elA = document.getElementById('prodKpiAvgConv');
        if (elQ) elQ.textContent = formatNumber(totalQty);
        if (elR) elR.textContent = '\u20B9' + fmtShort(totalRev);
        if (elC) elC.textContent = cats.length;
        if (elA) elA.textContent = avgConv.toFixed(2) + '%';

        // ----- Chart colours -----
        const CHART_COLORS = [
            '#f97316','#3b82f6','#10b981','#8b5cf6','#ec4899',
            '#f59e0b','#06b6d4','#ef4444','#84cc16','#6366f1'
        ];
        const top10 = cats.slice(0, 10);
        const labels = top10.map(c => c.name);
        const revData = top10.map(c => Math.round(c.pRev));
        const qtyData = top10.map(c => c.pQty);
        const convData = top10.map(c => parseFloat(c.osgConv.toFixed(2)));

        const chartDefaults = {
            responsive: true, maintainAspectRatio: false,
            plugins: { legend: { labels: { color: '#94a3b8', font: { size: 12 } } } }
        };

        // Destroy old charts
        ['prodRevChart', 'prodShareChart', 'prodConvChart'].forEach(id => {
            if (prodChartInstances[id]) { prodChartInstances[id].destroy(); delete prodChartInstances[id]; }
        });

        // Revenue Bar Chart
        const ctxRev = document.getElementById('prodRevChart');
        if (ctxRev) {
            prodChartInstances['prodRevChart'] = new Chart(ctxRev, {
                type: 'bar',
                data: {
                    labels,
                    datasets: [{
                        label: 'Revenue (\u20B9)',
                        data: revData,
                        backgroundColor: CHART_COLORS,
                        borderRadius: 6,
                        borderSkipped: false
                    }]
                },
                options: {
                    ...chartDefaults,
                    scales: {
                        x: { ticks: { color: '#94a3b8', font: { size: 11 } }, grid: { color: 'rgba(255,255,255,0.05)' } },
                        y: { ticks: { color: '#94a3b8', callback: v => '\u20B9' + fmtShort(v) }, grid: { color: 'rgba(255,255,255,0.07)' } }
                    },
                    plugins: { ...chartDefaults.plugins, tooltip: { callbacks: { label: ctx => '\u20B9' + formatNumber(ctx.parsed.y) } } }
                }
            });
        }

        // Doughnut Share Chart
        const ctxShare = document.getElementById('prodShareChart');
        if (ctxShare) {
            prodChartInstances['prodShareChart'] = new Chart(ctxShare, {
                type: 'doughnut',
                data: {
                    labels,
                    datasets: [{
                        data: qtyData,
                        backgroundColor: CHART_COLORS,
                        borderColor: '#1e293b',
                        borderWidth: 2,
                        hoverOffset: 8
                    }]
                },
                options: {
                    ...chartDefaults,
                    cutout: '62%',
                    plugins: {
                        legend: { position: 'right', labels: { color: '#94a3b8', boxWidth: 14, font: { size: 11 } } },
                        tooltip: { callbacks: { label: ctx => `${ctx.label}: ${formatNumber(ctx.parsed)} units` } }
                    }
                }
            });
        }

        // Conversion Bar Chart
        const ctxConv = document.getElementById('prodConvChart');
        if (ctxConv) {
            prodChartInstances['prodConvChart'] = new Chart(ctxConv, {
                type: 'bar',
                data: {
                    labels,
                    datasets: [{
                        label: 'OSG Conv%',
                        data: convData,
                        backgroundColor: convData.map(v => v >= 15 ? '#10b981' : v >= 10 ? '#f59e0b' : '#ef4444'),
                        borderRadius: 6,
                        borderSkipped: false
                    }]
                },
                options: {
                    ...chartDefaults,
                    scales: {
                        x: { ticks: { color: '#94a3b8', font: { size: 11 } }, grid: { color: 'rgba(255,255,255,0.05)' } },
                        y: { max: 100, ticks: { color: '#94a3b8', callback: v => v + '%' }, grid: { color: 'rgba(255,255,255,0.07)' } }
                    },
                    plugins: { ...chartDefaults.plugins, tooltip: { callbacks: { label: ctx => ctx.parsed.y.toFixed(2) + '%' } } }
                }
            });
        }

        // ----- Detail Table -----
        const tbody = document.getElementById('prodDetailTbody');
        if (!tbody) return;
        const fmtPct = (p) => {
            const color = p >= 15 ? 'color:#10b981;' : (p >= 10 ? 'color:#f59e0b;' : 'color:#ef4444;');
            return `<span style="font-weight:600;${color}">${p.toFixed(2)}%</span>`;
        };
        tbody.innerHTML = cats.map((c, i) => `
            <tr>
                <td style="font-weight:600; display:flex; align-items:center; gap:8px;">
                    <span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:${CHART_COLORS[i % 10]};"></span>
                    ${c.name}
                </td>
                <td class="text-right">${formatNumber(c.pQty)}</td>
                <td class="text-right">\u20B9${fmtShortHtml(c.pRev)}</td>
                <td class="text-right">${formatNumber(c.oQty)}</td>
                <td class="text-right">${fmtPct(c.osgConv)}</td>
                <td class="text-right">${formatNumber(c.aQty)}</td>
                <td class="text-right">${fmtPct(c.amcConv)}</td>
                <td class="text-right">\u20B9${fmtShortHtml(c.avgPrice)}</td>
            </tr>
        `).join('');
    }

    // ---- TRENDS DASHBOARD ----
    const trendChartInstances = {};

    async function renderTrendsDashboard() {
        const allData = await getAllMonthlyDataFromDB();

        // Destroy old charts
        ['trendConvChart', 'trendRevChart'].forEach(id => {
            if (trendChartInstances[id]) { trendChartInstances[id].destroy(); delete trendChartInstances[id]; }
        });

        if (!allData || allData.length === 0) {
            const ctxConv = document.getElementById('trendConvChart');
            const ctxRev = document.getElementById('trendRevChart');
            if (ctxConv) ctxConv.getContext('2d').clearRect(0, 0, ctxConv.width, ctxConv.height);
            if (ctxRev) ctxRev.getContext('2d').clearRect(0, 0, ctxRev.width, ctxRev.height);
            return;
        }

        // Sort by month ascending
        allData.sort((a, b) => a.month.localeCompare(b.month));

        const labels = allData.map(d => {
            const [y, m] = d.month.split('-');
            const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
            return `${months[parseInt(m)-1]} ${y}`;
        });
        const qtyConvData = allData.map(d => d.qtyConv || 0);
        const valConvData = allData.map(d => d.valConv || 0);
        const revData = allData.map(d => d.totalRev || 0);

        const chartDefaults = {
            responsive: true, maintainAspectRatio: false,
            plugins: { legend: { labels: { color: '#94a3b8', font: { size: 12 } } } },
            scales: {
                x: { ticks: { color: '#94a3b8', font: { size: 11 } }, grid: { color: 'rgba(255,255,255,0.05)' } },
                y: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.07)' } }
            }
        };

        // Conversion Trend (Dual Line)
        const ctxConv = document.getElementById('trendConvChart');
        if (ctxConv) {
            trendChartInstances['trendConvChart'] = new Chart(ctxConv, {
                type: 'line',
                data: {
                    labels,
                    datasets: [
                        {
                            label: 'Qty Conv%',
                            data: qtyConvData,
                            borderColor: '#f97316',
                            backgroundColor: 'rgba(249,115,22,0.1)',
                            tension: 0.4, fill: true, pointRadius: 5,
                            pointBackgroundColor: '#f97316', borderWidth: 2.5
                        },
                        {
                            label: 'Value Conv%',
                            data: valConvData,
                            borderColor: '#3b82f6',
                            backgroundColor: 'rgba(59,130,246,0.1)',
                            tension: 0.4, fill: true, pointRadius: 5,
                            pointBackgroundColor: '#3b82f6', borderWidth: 2.5
                        }
                    ]
                },
                options: {
                    ...chartDefaults,
                    scales: {
                        ...chartDefaults.scales,
                        y: { ...chartDefaults.scales.y, ticks: { ...chartDefaults.scales.y.ticks, callback: v => v + '%' } }
                    },
                    plugins: {
                        ...chartDefaults.plugins,
                        tooltip: { callbacks: { label: ctx => ctx.dataset.label + ': ' + ctx.parsed.y.toFixed(2) + '%' } }
                    }
                }
            });
        }

        // Revenue Trend (Area Chart)
        const ctxRev = document.getElementById('trendRevChart');
        if (ctxRev) {
            trendChartInstances['trendRevChart'] = new Chart(ctxRev, {
                type: 'line',
                data: {
                    labels,
                    datasets: [{
                        label: 'Total Revenue (\u20B9)',
                        data: revData,
                        borderColor: '#10b981',
                        backgroundColor: 'rgba(16,185,129,0.15)',
                        tension: 0.4, fill: true, pointRadius: 5,
                        pointBackgroundColor: '#10b981', borderWidth: 2.5
                    }]
                },
                options: {
                    ...chartDefaults,
                    scales: {
                        ...chartDefaults.scales,
                        y: { ...chartDefaults.scales.y, ticks: { ...chartDefaults.scales.y.ticks, callback: v => '\u20B9' + fmtShort(v) } }
                    },
                    plugins: {
                        ...chartDefaults.plugins,
                        tooltip: { callbacks: { label: ctx => '\u20B9' + formatNumber(ctx.parsed.y) } }
                    }
                }
            });
        }
    }

    // Load trends on startup
    renderTrendsDashboard();

    function renderPincodeDashboard() {
        const tbody = document.querySelector('#pincodeTable tbody');
        const noData = document.getElementById('pincodeNoData');
        const wrapper = document.getElementById('pincodeTableWrapper');
        if (!tbody || !noData || !wrapper) return;

        // Same filtering logic as the main dashboard
        const dataProduct = filteredProduct;
        if (!dataProduct || dataProduct.length === 0) {
            tbody.innerHTML = '';
            noData.style.display = 'flex';
            wrapper.style.display = 'none';
            return;
        }

        noData.style.display = 'none';
        wrapper.style.display = 'block';

        const invMeta = {};
        dataProduct.forEach(r => {
            if (r.invoice) invMeta[r.invoice] = { branch: r.branch, bdm: r.bdm || 'Unknown', pincode: r.pincode || 'Unknown' };
        });

        const pincodeGroups = {};
        dataProduct.forEach(r => {
            const pKey = r.pincode || 'Unknown';
            if (!pincodeGroups[pKey]) pincodeGroups[pKey] = { pincode: pKey, pRows: [], oRows: [], aRows: [], sRows: [] };
            pincodeGroups[pKey].pRows.push(r);
        });

        filteredOSG.forEach(r => {
            if (r.invoice && invMeta[r.invoice]) {
                const pKey = invMeta[r.invoice].pincode;
                if (pincodeGroups[pKey]) pincodeGroups[pKey].oRows.push(r);
            }
        });
        filteredAMC.forEach(r => {
            // Use AMC row's own pincode if available, else fallback to invoice lookup
            let pKey;
            if (r.pincode) {
                pKey = r.pincode;
            } else if (r.invoice && invMeta[r.invoice]) {
                pKey = invMeta[r.invoice].pincode;
            } else {
                pKey = 'Unknown';
            }
            if (!pincodeGroups[pKey]) pincodeGroups[pKey] = { pincode: pKey, pRows: [], oRows: [], aRows: [], sRows: [] };
            pincodeGroups[pKey].aRows.push(r);
        });
        filteredSamsung.forEach(r => {
            if (r.invoice && invMeta[r.invoice]) {
                const pKey = invMeta[r.invoice].pincode;
                if (pincodeGroups[pKey]) pincodeGroups[pKey].sRows.push(r);
            }
        });

        const samsungAllowedCats = ['AC', 'MICROWAVE OVEN', 'REFRIGERATOR', 'WASHING MACHINE', 'TV'];
        const stats = Object.values(pincodeGroups).sort((a, b) => {
            // Default sort by Total Product Qty descending
            const aQty = a.pRows.reduce((s, r) => s + (r.qty || 0), 0);
            const bQty = b.pRows.reduce((s, r) => s + (r.qty || 0), 0);
            return bQty - aQty;
        });

        let html = '';
        stats.forEach(grp => {
            const pQty = grp.pRows.reduce((s, r) => s + (r.qty || 0), 0);
            if (pQty === 0) return; // Skip empty pincodes

            const pRev = grp.pRows.reduce((s, r) => s + (r.revenue || r.soldPrice || 0), 0);
            const oQty = grp.oRows.reduce((s, r) => s + (r.qty || 0), 0);
            const aQty = grp.aRows.reduce((s, r) => s + (r.qty || 0), 0);
            const sQty = grp.sRows.reduce((s, r) => s + (r.qty || 0), 0);

            const lgPQty = grp.pRows.filter(r => r.brand && r.brand.toUpperCase().includes('LG')).reduce((s, r) => s + (r.qty || 0), 0);
            const samPQty = grp.pRows.filter(r => r.brand && r.brand.toUpperCase().includes('SAMSUNG') && samsungAllowedCats.includes((r.product || '').toUpperCase().trim())).reduce((s, r) => s + (r.qty || 0), 0);

            const osgQtyC = pQty > 0 ? (oQty / pQty) * 100 : 0;
            const amcQtyC = lgPQty > 0 ? (aQty / lgPQty) * 100 : 0;
            const samQtyC = samPQty > 0 ? (sQty / samPQty) * 100 : 0;

            const fmtNum = (n) => `<span class="kpi-value">${formatNumber(n)}</span>`;
            const fmtPct = (p) => {
                const color = p >= 15 ? 'color:#10b981;' : (p >= 10 ? 'color:#f59e0b;' : 'color:#ef4444;');
                return `<span style="font-weight:600; ${color}">${p.toFixed(2)}%</span>`;
            };

            html += `
                <tr>
                    <td style="font-weight:600; color:var(--text-main);">${grp.pincode}</td>
                    <td id="area-${grp.pincode}" style="font-size:0.85rem; color:var(--text-muted);">
                        ${pincodeAreaCache[grp.pincode] ? pincodeAreaCache[grp.pincode] : 'Loading...'}
                    </td>
                    <td class="text-right">${fmtNum(pQty)}</td>
                    <td class="text-right">\u20B9${fmtShortHtml(pRev)}</td>
                    <td class="text-right">${fmtNum(oQty)}</td>
                    <td class="text-right">${fmtPct(osgQtyC)}</td>
                    <td class="text-right">${fmtNum(aQty)}</td>
                    <td class="text-right">${fmtPct(amcQtyC)}</td>
                    <td class="text-right">${fmtNum(sQty)}</td>
                    <td class="text-right">${fmtPct(samQtyC)}</td>
                </tr>
            `;

            // Trigger async fetch if we don't have it cached
            if (!pincodeAreaCache[grp.pincode]) {
                fetchAreaForPincode(grp.pincode);
            }
        });

        tbody.innerHTML = html;
    }

    async function fetchAreaForPincode(pincode) {
        if (!pincode || pincode === 'Unknown' || String(pincode).length < 6) {
            pincodeAreaCache[pincode] = '-';
            updateAreaUI(pincode, '-');
            return;
        }

        try {
            const res = await fetch(`https://api.postalpincode.in/pincode/${pincode}`);
            const data = await res.json();
            if (data && data[0] && data[0].Status === 'Success' && data[0].PostOffice && data[0].PostOffice.length > 0) {
                // Get the first area name, or join a couple if desired. Just taking the primary PostOffice Name
                const areaName = data[0].PostOffice[0].Name;
                const district = data[0].PostOffice[0].District;
                const fullName = `${areaName}, ${district}`;
                pincodeAreaCache[pincode] = fullName;
                updateAreaUI(pincode, fullName);
            } else {
                pincodeAreaCache[pincode] = 'Not Found';
                updateAreaUI(pincode, 'Not Found');
            }
        } catch (err) {
            console.error('Error fetching area for pincode', pincode, err);
            updateAreaUI(pincode, 'Error');
        }
    }

    function updateAreaUI(pincode, areaName) {
        const td = document.getElementById(`area-${pincode}`);
        if (td) {
            td.textContent = areaName;
        }
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
            // Build invoice ÃƒÂ¢Ã¢â‚¬Â Ã¢â‚¬â„¢ key value lookup from filtered product data
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
                // Skip OSG rows that can't be attributed ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â no "Unknown"
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
                <td class="number-cell">${fmtShortHtml(e.pSold)}</td>
                <td class="number-cell">${fmtShortHtml(e.oSold)}</td>
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
        // Group by branch ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â show value and qty conversion
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
                <td class="number-cell">${fmtShortHtml(e.pSold)}</td>
                <td class="number-cell">${fmtShortHtml(e.oSold)}</td>
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
                <td class="number-cell">${fmtShortHtml(r.soldPrice)}</td>
                <td class="number-cell">${fmtShortHtml(r.taxableVal)}</td>
                <td class="number-cell">${fmtShortHtml(r.tax)}</td>
                <td class="number-cell">${r.qty}</td>
                <td class="number-cell ${cls}">${fmtShortHtml(r.profit)}</td>
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

        // Map OSG ÃƒÂ¢Ã¢â‚¬Â Ã¢â‚¬â„¢ RBM via invoice lookup (same as renderConvTable)
        const invoiceToRBM = {};
        filteredProduct.forEach(r => { if (r.invoice && r.rbm) invoiceToRBM[r.invoice] = r.rbm; });
        const osgByRBM = {};
        filteredOSG.forEach(r => {
            const rbm = r.invoice ? (invoiceToRBM[r.invoice] || null) : null;
            if (!rbm) return; // skip unattributable rows ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â no "Unknown"
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

    // Generic EXPORT CSV listener removed to prevent undefined_report.csv bugs

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
    (function () {
        var btnResetEl = document.getElementById('btnReset');
        if (!btnResetEl) return;
        btnResetEl.addEventListener('click', function () {
            productData = []; osgData = []; amcData = []; samsungData = []; allData = [];
            filteredProduct = []; filteredOSG = []; filteredAMC = []; filteredSamsung = []; filteredAll = [];
            const pSt = document.getElementById('productStatus');
            const oSt = document.getElementById('osgStatus');
            const aSt = document.getElementById('amcStatus');
            const sSt = document.getElementById('samsungStatus');
            if (pSt) { pSt.className = 'upload-status'; pSt.innerHTML = ''; }
            if (oSt) { oSt.className = 'upload-status'; oSt.innerHTML = ''; }
            if (aSt) { aSt.className = 'upload-status'; aSt.innerHTML = ''; }
            if (sSt) { sSt.className = 'upload-status'; sSt.innerHTML = ''; }
            const fcb = document.getElementById('fileCountBadge');
            const bGen = document.getElementById('btnGenerate');
            const bShr = document.getElementById('btnShare');
            if (fcb) fcb.style.display = 'none';
            btnResetEl.style.display = 'none';
            if (bShr) bShr.style.display = 'none';
            if (bGen) bGen.disabled = true;
            [filterRBM, filterBranch, filterBDM, filterStaff].forEach(sel => {
                if (sel && sel.innerHTML !== undefined) {
                    const textContent = (sel.options && sel.options.length > 0) ? sel.options[0].textContent : 'All';
                    sel.innerHTML = `<option value="">${textContent || 'All'}</option>`;
                }
            });
            if (filterProduct) filterProduct.value = '';
            if (filterBrand) filterBrand.innerHTML = '<option value="">All Brands</option>';
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
    })();

    // ---- LOW CONV STAFF PAGE ----
    $('btnLCRefresh').addEventListener('click', renderLowConvPage);
    $('btnLCExport').addEventListener('click', exportLowConvCSV);
    if ($('btnLCExportDetailed')) $('btnLCExportDetailed').addEventListener('click', exportDetailedLowConvCSV);

    // Also refresh whenever user navigates to the page
    document.querySelector('[data-section="lowconv-section"]').addEventListener('click', () => {
        setTimeout(renderLowConvPage, 50);
    });

    // ---- AI GOD MODE DATA EXPORTER ----
    window.getGodModeContextData = function() {
        var catMap = {};
        var branchMap = {};
        var brandMap = {};
        var rbmMap = {};
        var bdmMap = {};
        
        productData.forEach(r => {
            var c = r.category || 'Unknown';
            var b = r.branch || 'Unknown';
            var br = r.brand || 'Unknown';
            var rbm = r.rbm || 'Unknown';
            var bdm = r.bdm || 'Unknown';

            if (!catMap[c]) catMap[c] = { pQty: 0, oQty: 0 };
            if (!branchMap[b]) branchMap[b] = { pQty: 0, oQty: 0 };
            if (!brandMap[br]) brandMap[br] = { pQty: 0, oQty: 0 };
            if (!rbmMap[rbm]) rbmMap[rbm] = { pQty: 0, oQty: 0 };
            if (!bdmMap[bdm]) bdmMap[bdm] = { pQty: 0, oQty: 0 };

            catMap[c].pQty += r.qty || 0;
            branchMap[b].pQty += r.qty || 0;
            brandMap[br].pQty += r.qty || 0;
            rbmMap[rbm].pQty += r.qty || 0;
            bdmMap[bdm].pQty += r.qty || 0;
        });

        osgData.forEach(r => {
            var c = r.category || 'Unknown';
            var b = r.branch || 'Unknown';
            var br = r.brand || 'Unknown';
            var rbm = r.rbm || 'Unknown';
            var bdm = r.bdm || 'Unknown';

            if (!catMap[c]) catMap[c] = { pQty: 0, oQty: 0 };
            if (!branchMap[b]) branchMap[b] = { pQty: 0, oQty: 0 };
            if (!brandMap[br]) brandMap[br] = { pQty: 0, oQty: 0 };
            if (!rbmMap[rbm]) rbmMap[rbm] = { pQty: 0, oQty: 0 };
            if (!bdmMap[bdm]) bdmMap[bdm] = { pQty: 0, oQty: 0 };

            catMap[c].oQty += r.qty || 0;
            branchMap[b].oQty += r.qty || 0;
            brandMap[br].oQty += r.qty || 0;
            rbmMap[rbm].oQty += r.qty || 0;
            bdmMap[bdm].oQty += r.qty || 0;
        });

        var staffStats = window.portalStaffStats;
        if (!staffStats || staffStats.length === 0) {
            try {
                staffStats = buildStaffStats();
            } catch(e) {
                staffStats = [];
            }
        }

        var missingFollowups = typeof missedUnique !== 'undefined' ? missedUnique.length : 0;

        var callerStats = {};
        if (typeof CO_CALLERS !== 'undefined') {
            CO_CALLERS.forEach(c => callerStats[c.name] = { totalCalls: 0, connected: 0, disconnected: 0, interested: 0, bought: 0 });
        }
        if (typeof coStatusMap !== 'undefined') {
            Object.values(coStatusMap).forEach(st => {
                if (st.calledBy) {
                    if (!callerStats[st.calledBy]) callerStats[st.calledBy] = { totalCalls: 0, connected: 0, disconnected: 0, interested: 0, bought: 0 };
                    callerStats[st.calledBy].totalCalls++;
                    if (st.callStatus === 'connected') callerStats[st.calledBy].connected++;
                    if (st.callStatus === 'disconnected') callerStats[st.calledBy].disconnected++;
                    if (st.interest === 'interested') callerStats[st.calledBy].interested++;
                    if (st.interest === 'bought') callerStats[st.calledBy].bought++;
                }
            });
        }

        var sortedStaff = staffStats.filter(s => s.pQty > 0).sort((a, b) => b.qtyConv - a.qtyConv);
        var bestStaff = sortedStaff.slice(0, 15);
        var worstStaff = sortedStaff.slice(-15).reverse();

        return {
            Brands: Object.keys(brandMap).map(k => ({ Brand: k, ProductsSold: brandMap[k].pQty, OsgSold: brandMap[k].oQty, ConvPercent: (brandMap[k].pQty>0 ? (brandMap[k].oQty/brandMap[k].pQty*100).toFixed(1) : 0) })),
            Categories: Object.keys(catMap).map(k => ({ Category: k, ProductsSold: catMap[k].pQty, OsgSold: catMap[k].oQty, ConvPercent: (catMap[k].pQty>0 ? (catMap[k].oQty/catMap[k].pQty*100).toFixed(1) : 0) })),
            Branches: Object.keys(branchMap).map(k => ({ Branch: k, ProductsSold: branchMap[k].pQty, OsgSold: branchMap[k].oQty, ConvPercent: (branchMap[k].pQty>0 ? (branchMap[k].oQty/branchMap[k].pQty*100).toFixed(1) : 0) })),
            RBMs: Object.keys(rbmMap).map(k => ({ RBM: k, ProductsSold: rbmMap[k].pQty, OsgSold: rbmMap[k].oQty, ConvPercent: (rbmMap[k].pQty>0 ? (rbmMap[k].oQty/rbmMap[k].pQty*100).toFixed(1) : 0) })),
            BDMs: Object.keys(bdmMap).map(k => ({ BDM: k, ProductsSold: bdmMap[k].pQty, OsgSold: bdmMap[k].oQty, ConvPercent: (bdmMap[k].pQty>0 ? (bdmMap[k].oQty/bdmMap[k].pQty*100).toFixed(1) : 0) })),
            Top15_Best_Staff: bestStaff.map(s => ({ StaffName: s.name, Branch: s.branch, ProductsSold: s.pQty, Total_Revenue: s.pRev, OsgSold: s.oQty, OSG_Revenue: s.oRev, ConvPercent: s.qtyConv.toFixed(1) })),
            Top15_Worst_Staff: worstStaff.map(s => ({ StaffName: s.name, Branch: s.branch, ProductsSold: s.pQty, Total_Revenue: s.pRev, OsgSold: s.oQty, OSG_Revenue: s.oRev, ConvPercent: s.qtyConv.toFixed(1) })),
            MissingCRM_Count: missingFollowups,
            CRM_Callers: callerStats
        };
    };

    window.getLcSelectedProducts = function() {
        if (!document.getElementById('lcProductDropdown')) return null;
        const cbs = document.querySelectorAll('.lc-prod-cb');
        const selected = [];
        cbs.forEach(cb => { if (cb.checked) selected.push(cb.value); });
        const allCb = document.querySelector('#lcProductDropdown input[value="ALL"]');
        if (selected.length === cbs.length || (allCb && allCb.checked)) return null; // All products selected
        return selected;
    };

    window.toggleAllLcProducts = function(allCb) {
        document.querySelectorAll('.lc-prod-cb').forEach(cb => cb.checked = allCb.checked);
        window.updateLcProductLabel();
    };

    window.updateLcProductLabel = function() {
        const cbs = Array.from(document.querySelectorAll('.lc-prod-cb'));
        const allCb = document.querySelector('#lcProductDropdown input[value="ALL"]');
        const checked = cbs.filter(cb => cb.checked);
        const label = document.getElementById('lcProductLabel');
        if (checked.length === cbs.length) {
            if (allCb) allCb.checked = true;
            label.textContent = "All Products";
        } else if (checked.length === 0) {
            if (allCb) allCb.checked = false;
            label.textContent = "None Selected";
        } else {
            if (allCb) allCb.checked = false;
            label.textContent = checked.length + " Selected";
        }
    };

    window.getLcSelectedBranches = function() {
        if (!document.getElementById('lcBranchDropdown')) return null;
        const cbs = document.querySelectorAll('.lc-branch-cb');
        const selected = [];
        cbs.forEach(cb => { if (cb.checked) selected.push(cb.value); });
        const allCb = document.querySelector('#lcBranchDropdown input[value="ALL"]');
        if (selected.length === cbs.length || (allCb && allCb.checked)) return null;
        return selected;
    };

    window.toggleAllLcBranches = function(allCb) {
        document.querySelectorAll('.lc-branch-cb').forEach(cb => cb.checked = allCb.checked);
        window.updateLcBranchLabel();
    };

    window.updateLcBranchLabel = function() {
        const cbs = Array.from(document.querySelectorAll('.lc-branch-cb'));
        const allCb = document.querySelector('#lcBranchDropdown input[value="ALL"]');
        const checked = cbs.filter(cb => cb.checked);
        const label = document.getElementById('lcBranchLabel');
        if (!label) return;
        if (checked.length === cbs.length) {
            if (allCb) allCb.checked = true;
            label.textContent = "All Branches";
        } else if (checked.length === 0) {
            if (allCb) allCb.checked = false;
            label.textContent = "None Selected";
        } else {
            if (allCb) allCb.checked = false;
            label.textContent = checked.length + " Selected";
        }
    };

    window.getTcSelectedBranches = function() {
        if (!document.getElementById('tcBranchDropdown')) return null;
        const cbs = document.querySelectorAll('.tc-branch-cb');
        const selected = [];
        cbs.forEach(cb => { if (cb.checked) selected.push(cb.value); });
        const allCb = document.querySelector('#tcBranchDropdown input[value="ALL"]');
        if (selected.length === cbs.length || (allCb && allCb.checked)) return null;
        return selected;
    };

    window.toggleAllTcBranches = function(allCb) {
        document.querySelectorAll('.tc-branch-cb').forEach(cb => cb.checked = allCb.checked);
        window.updateTcBranchLabel();
    };

    window.updateTcBranchLabel = function() {
        const cbs = Array.from(document.querySelectorAll('.tc-branch-cb'));
        const allCb = document.querySelector('#tcBranchDropdown input[value="ALL"]');
        const checked = cbs.filter(cb => cb.checked);
        const label = document.getElementById('tcBranchLabel');
        if (!label) return;
        if (checked.length === cbs.length) {
            if (allCb) allCb.checked = true;
            label.textContent = "All Branches";
        } else if (checked.length === 0) {
            if (allCb) allCb.checked = false;
            label.textContent = "None Selected";
        } else {
            if (allCb) allCb.checked = false;
            label.textContent = checked.length + " Selected";
        }
    };

    // Close dropdown on click outside
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
    });

    // ---- LOW CONV STAFF LOGIC ----
    function buildStaffStats(selectedProducts = null) {
        // Build invoice -> {staff, product} lookup from ALL product data
        const invoiceData = {};
        productData.forEach(r => { if (r.invoice) invoiceData[r.invoice] = { staff: r.staff || 'Unknown', product: r.category }; });

        // Group product data by staff
        const pByStaff = {};
        productData.forEach(r => {
            if (selectedProducts && !selectedProducts.includes(r.category)) return;
            const s = r.staff || 'Unknown';
            if (!pByStaff[s]) pByStaff[s] = { branch: r.branch, rbm: r.rbm, bdm: r.bdm, rows: [] };
            pByStaff[s].rows.push(r);
        });

        // Group OSG data by staff via invoice mapping
        const oByStaff = {};
        osgData.forEach(r => {
            const inv = r.invoice ? invoiceData[r.invoice] : null;
            const s = inv ? inv.staff : null;
            const pName = inv ? inv.product : null;
            if (!s) return;
            if (selectedProducts && !selectedProducts.includes(pName)) return;
            if (!oByStaff[s]) oByStaff[s] = [];
            oByStaff[s].push(r);
        });

        const allStaff = new Set([...Object.keys(pByStaff), ...Object.keys(oByStaff)]);
        allStaff.delete('Unknown');

        const finalStats = Array.from(allStaff).map(name => {
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
        window.portalStaffStats = finalStats;
        return finalStats;
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
        const minOsgQty = parseInt($('lcMinOsgQty').value) || 0;
        const selectedBranches = window.getLcSelectedBranches();
        const selRBM = $('lcRBM').value;
        const selBDM = $('lcBDM').value;

        const selectedProducts = window.getLcSelectedProducts();
        const allStats = buildStaffStats(selectedProducts);

        // Populate dropdowns (preserve selection)
        const branchSet = [...new Set(allStats.map(s => s.branch).filter(Boolean))].sort();
        const rbmSet = [...new Set(allStats.map(s => s.rbm).filter(Boolean))].sort();
        const bdmSet = [...new Set(allStats.map(s => s.bdm).filter(Boolean))].sort();

        const branchDrop = $('lcBranchDropdown');
        const rbmEl = $('lcRBM');
        const bdmEl = $('lcBDM');
        const prevRBM = selRBM;
        const prevBDM = selBDM;

        if (branchDrop) {
            let bHtml = `<label style="display: flex; align-items: center; gap: 8px; padding: 4px 0; cursor: pointer; color: var(--text-primary); text-transform: none; font-weight: 500; font-size: 0.85rem; margin:0;"><input type="checkbox" value="ALL" ${!selectedBranches ? 'checked' : ''} onchange="window.toggleAllLcBranches(this)"> <strong>All Branches</strong></label><hr style="margin: 4px 0; border: none; border-top: 1px solid var(--border);">`;
            bHtml += branchSet.map(b => `<label style="display: flex; align-items: center; gap: 8px; padding: 4px 0; cursor: pointer; color: var(--text-primary); text-transform: none; font-weight: 500; font-size: 0.85rem; margin:0;"><input type="checkbox" value="${b}" class="lc-branch-cb" ${!selectedBranches || selectedBranches.includes(b) ? 'checked' : ''} onchange="window.updateLcBranchLabel()"> ${b}</label>`).join('');
            branchDrop.innerHTML = bHtml;
            window.updateLcBranchLabel();
        }
        rbmEl.innerHTML = '<option value="">All RBMs</option>' +
            rbmSet.map(r => `<option value="${r}" ${r === prevRBM ? 'selected' : ''}>${r}</option>`).join('');
        bdmEl.innerHTML = '<option value="">All BDMs</option>' +
            bdmSet.map(b => `<option value="${b}" ${b === prevBDM ? 'selected' : ''}>${b}</option>`).join('');

        // Filter: minQty, maxConv, minOsgQty, optional Branch, optional RBM, optional BDM
        const filtered = allStats
            .filter(s => s.pQty >= minQty && s.qtyConv <= maxConv)
            .filter(s => s.oQty >= minOsgQty)
            .filter(s => !selectedBranches || selectedBranches.includes(s.branch))
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
            <div class="lc-kpi"><span class="lc-kpi-label">Total Product Revenue</span><span class="lc-kpi-val">${fmtShortHtml(totalPRev)}</span></div>
        `;

        if (filtered.length === 0) {
            $('lcTableWrapper').innerHTML = noDataHTML(`No staff found with â‰¥${minQty} product qty and ÃƒÂ¢Ã¢â‚¬Â°Ã‚Â¤${maxConv}% qty conversion.`);
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
                <td class="number-cell">${fmtShortHtml(e.pRev)}</td>
            </tr>`;
        });

        html += '</tbody></table>';
        $('lcTableWrapper').innerHTML = html;
    }

    function exportLowConvCSV() {
        if (productData.length === 0) return;
        const minQty = parseFloat($('lcMinQty').value) || 0;
        const maxConv = parseFloat($('lcMaxConv').value);
        const minOsgQty = parseInt($('lcMinOsgQty').value) || 0;
        const selectedBranches = window.getLcSelectedBranches();
        const selRBM = $('lcRBM').value;
        const selBDM = $('lcBDM').value;
        const selectedProducts = window.getLcSelectedProducts();
        
        const allStats = buildStaffStats(selectedProducts);
        const filteredStaff = allStats
            .filter(s => s.pQty >= minQty && s.qtyConv <= maxConv)
            .filter(s => s.oQty >= minOsgQty)
            .filter(s => !selectedBranches || selectedBranches.includes(s.branch))
            .filter(s => !selRBM || s.rbm === selRBM)
            .filter(s => !selBDM || s.bdm === selBDM);
            
        if (filteredStaff.length === 0) {
            alert("No staff match the current filters.");
            return;
        }

        // Apply Top Staff Limit ONLY for Export CSV
        const exportLimitEl = $('lcTopStaffExportLimit');
        const topLimit = exportLimitEl ? (parseInt(exportLimitEl.value) || 15) : 15;
        filteredStaff.sort((a, b) => a.qtyConv - b.qtyConv || b.pQty - a.pQty);
        const topStaff = filteredStaff.slice(0, topLimit);
        
        const validStaffNames = new Set(topStaff.map(s => s.name));
        const detailedMap = {};
        
        productData.forEach(r => {
            const s = r.staff || 'Unknown';
            if (!validStaffNames.has(s)) return;
            if (selectedProducts && !selectedProducts.includes(r.category)) return;
            
            const cat = r.category || 'Unknown';
            const key = s + '|||' + cat;
            
            if (!detailedMap[key]) {
                detailedMap[key] = {
                    branch: r.branch || 'Unknown',
                    rbm: r.rbm || 'Unknown',
                    bdm: r.bdm || 'Unknown',
                    staff: s,
                    product: cat,
                    pQty: 0,
                    oQty: 0,
                    lgOsgQty: 0,
                    samsungOsgQty: 0
                };
            }
            detailedMap[key].pQty += (r.qty || 0);
        });
        
        const invoiceData = {};
        productData.forEach(r => { 
            if (r.invoice) invoiceData[r.invoice] = { staff: r.staff || 'Unknown', product: r.category }; 
        });
        
        osgData.forEach(r => {
            const inv = r.invoice ? invoiceData[r.invoice] : null;
            if (!inv) return;
            
            const s = inv.staff;
            const cat = inv.product;
            if (!validStaffNames.has(s)) return;
            if (selectedProducts && !selectedProducts.includes(cat)) return;
            
            const key = s + '|||' + cat;
            if (detailedMap[key]) {
                const qty = r.qty || 0;
                detailedMap[key].oQty += qty;
                
                const osgBrand = (r.brand || '').toUpperCase();
                if (osgBrand === 'LG') detailedMap[key].lgOsgQty += qty;
                if (osgBrand === 'SAMSUNG') detailedMap[key].samsungOsgQty += qty;
            }
        });

        amcData.forEach(r => {
            const inv = r.invoice ? invoiceData[r.invoice] : null;
            if (!inv) return;
            const key = inv.staff + '|||' + inv.product;
            if (detailedMap[key]) {
                const qty = r.qty || 0;
                detailedMap[key].lgOsgQty += qty;
            }
        });
        
        samsungData.forEach(r => {
            const inv = r.invoice ? invoiceData[r.invoice] : null;
            if (!inv) return;
            const key = inv.staff + '|||' + inv.product;
            if (detailedMap[key]) {
                const qty = r.qty || 0;
                detailedMap[key].samsungOsgQty += qty;
            }
        });
        
        const detailedRows = Object.values(detailedMap);
        
        // Compute Totals per Staff
        const staffTotals = {};
        detailedRows.forEach(r => {
            if (!staffTotals[r.staff]) {
                staffTotals[r.staff] = {
                    branch: r.branch, rbm: r.rbm, bdm: r.bdm, staff: r.staff, product: 'TOTAL',
                    pQty: 0, oQty: 0, lgOsgQty: 0, samsungOsgQty: 0
                };
            }
            staffTotals[r.staff].pQty += r.pQty;
            staffTotals[r.staff].oQty += r.oQty;
            staffTotals[r.staff].lgOsgQty += r.lgOsgQty;
            staffTotals[r.staff].samsungOsgQty += r.samsungOsgQty;
        });

        // Add TOTAL rows to detailedRows
        Object.values(staffTotals).forEach(tot => detailedRows.push(tot));

        detailedRows.sort((a, b) => {
            if (a.rbm !== b.rbm) return a.rbm.localeCompare(b.rbm);
            if (a.branch !== b.branch) return a.branch.localeCompare(b.branch);
            if (a.staff !== b.staff) return a.staff.localeCompare(b.staff);
            // Ensure TOTAL comes first for the staff
            if (a.product === 'TOTAL') return -1;
            if (b.product === 'TOTAL') return 1;
            return a.product.localeCompare(b.product);
        });
        
        const hdr = ['BRANCH', 'RBM', 'BDM', 'Staff', 'Product', 'Product Qty', 'OSG QTY', 'LG', 'SAMSUNG'];
        const data = detailedRows.map(r => [
            r.branch, r.rbm, r.bdm, r.staff, r.product, r.pQty, r.oQty, r.lgOsgQty, r.samsungOsgQty
        ]);
        
        exportToStyledExcel(data, hdr, 'low_conv_staff.xlsx', 'Low Conversion Staff');
    }

    function exportDetailedLowConvCSV() {
        if (productData.length === 0) return;
        const minQty = parseFloat($('lcMinQty').value) || 0;
        const maxConv = parseFloat($('lcMaxConv').value);
        const minOsgQty = parseInt($('lcMinOsgQty').value) || 0;
        const selectedBranches = window.getLcSelectedBranches();
        const selRBM = $('lcRBM').value;
        const selBDM = $('lcBDM').value;
        const selectedProducts = window.getLcSelectedProducts();
        
        const allStats = buildStaffStats(selectedProducts);
        const filteredStaff = allStats
            .filter(s => s.pQty >= minQty && s.qtyConv <= maxConv)
            .filter(s => s.oQty >= minOsgQty)
            .filter(s => !selectedBranches || selectedBranches.includes(s.branch))
            .filter(s => !selRBM || s.rbm === selRBM)
            .filter(s => !selBDM || s.bdm === selBDM);
            
        if (filteredStaff.length === 0) {
            alert("No staff match the current filters.");
            return;
        }
        
        const validStaffNames = new Set(filteredStaff.map(s => s.name));
        const detailedMap = {};
        
        productData.forEach(r => {
            const s = r.staff || 'Unknown';
            if (!validStaffNames.has(s)) return;
            if (selectedProducts && !selectedProducts.includes(r.category)) return;
            
            const cat = r.category || 'Unknown';
            const key = s + '|||' + cat;
            
            if (!detailedMap[key]) {
                detailedMap[key] = {
                    branch: r.branch || 'Unknown',
                    rbm: r.rbm || 'Unknown',
                    bdm: r.bdm || 'Unknown',
                    staff: s,
                    product: cat,
                    pQty: 0,
                    oQty: 0,
                    lgOsgQty: 0,
                    samsungOsgQty: 0
                };
            }
            detailedMap[key].pQty += (r.qty || 0);
        });
        
        const invoiceData = {};
        productData.forEach(r => { 
            if (r.invoice) invoiceData[r.invoice] = { staff: r.staff || 'Unknown', product: r.category }; 
        });
        
        osgData.forEach(r => {
            const inv = r.invoice ? invoiceData[r.invoice] : null;
            if (!inv) return;
            
            const s = inv.staff;
            const cat = inv.product;
            if (!validStaffNames.has(s)) return;
            if (selectedProducts && !selectedProducts.includes(cat)) return;
            
            const key = s + '|||' + cat;
            if (detailedMap[key]) {
                const qty = r.qty || 0;
                detailedMap[key].oQty += qty;
                
                const osgBrand = (r.brand || '').toUpperCase();
                if (osgBrand === 'LG') detailedMap[key].lgOsgQty += qty;
                if (osgBrand === 'SAMSUNG') detailedMap[key].samsungOsgQty += qty;
            }
        });

        amcData.forEach(r => {
            const inv = r.invoice ? invoiceData[r.invoice] : null;
            if (!inv) return;
            const key = inv.staff + '|||' + inv.product;
            if (detailedMap[key]) {
                const qty = r.qty || 0;
                detailedMap[key].lgOsgQty += qty;
            }
        });
        
        samsungData.forEach(r => {
            const inv = r.invoice ? invoiceData[r.invoice] : null;
            if (!inv) return;
            const key = inv.staff + '|||' + inv.product;
            if (detailedMap[key]) {
                const qty = r.qty || 0;
                detailedMap[key].samsungOsgQty += qty;
            }
        });
        
        const detailedRows = Object.values(detailedMap);
        
        // Compute Totals per Staff
        const staffTotals = {};
        detailedRows.forEach(r => {
            if (!staffTotals[r.staff]) {
                staffTotals[r.staff] = {
                    branch: r.branch, rbm: r.rbm, bdm: r.bdm, staff: r.staff, product: 'TOTAL',
                    pQty: 0, oQty: 0, lgOsgQty: 0, samsungOsgQty: 0
                };
            }
            staffTotals[r.staff].pQty += r.pQty;
            staffTotals[r.staff].oQty += r.oQty;
            staffTotals[r.staff].lgOsgQty += r.lgOsgQty;
            staffTotals[r.staff].samsungOsgQty += r.samsungOsgQty;
        });

        // Add TOTAL rows to detailedRows
        Object.values(staffTotals).forEach(tot => detailedRows.push(tot));

        detailedRows.sort((a, b) => {
            if (a.rbm !== b.rbm) return a.rbm.localeCompare(b.rbm);
            if (a.branch !== b.branch) return a.branch.localeCompare(b.branch);
            if (a.staff !== b.staff) return a.staff.localeCompare(b.staff);
            // Ensure TOTAL comes first for the staff
            if (a.product === 'TOTAL') return -1;
            if (b.product === 'TOTAL') return 1;
            return a.product.localeCompare(b.product);
        });
        
        const hdr = ['BRANCH', 'RBM', 'BDM', 'Staff', 'Product', 'Product Qty', 'OSG QTY', 'LG', 'SAMSUNG'];
        const data = detailedRows.map(r => [
            r.branch, r.rbm, r.bdm, r.staff, r.product, r.pQty, r.oQty, r.lgOsgQty, r.samsungOsgQty
        ]);
        
        exportToStyledExcel(data, hdr, 'detailed_low_conv_staff.xlsx', 'Detailed Low Conv Staff');
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
        const selectedBranches = window.getTcSelectedBranches();
        const selRBM = $('tcRBM').value;
        const selBDM = $('tcBDM').value;

        const allStats = buildStaffStats();

        // Populate RBM and BDM dropdowns (preserve selection)
        const branchSet = [...new Set(allStats.map(s => s.branch).filter(Boolean))].sort();
        const rbmSet = [...new Set(allStats.map(s => s.rbm).filter(Boolean))].sort();
        const bdmSet = [...new Set(allStats.map(s => s.bdm).filter(Boolean))].sort();
        const branchDrop = $('tcBranchDropdown');
        if (branchDrop) {
            let bHtml = `<label style="display: flex; align-items: center; gap: 8px; padding: 4px 0; cursor: pointer; color: var(--text-primary); text-transform: none; font-weight: 500; font-size: 0.85rem; margin:0;"><input type="checkbox" value="ALL" ${!selectedBranches ? 'checked' : ''} onchange="window.toggleAllTcBranches(this)"> <strong>All Branches</strong></label><hr style="margin: 4px 0; border: none; border-top: 1px solid var(--border);">`;
            bHtml += branchSet.map(b => `<label style="display: flex; align-items: center; gap: 8px; padding: 4px 0; cursor: pointer; color: var(--text-primary); text-transform: none; font-weight: 500; font-size: 0.85rem; margin:0;"><input type="checkbox" value="${b}" class="tc-branch-cb" ${!selectedBranches || selectedBranches.includes(b) ? 'checked' : ''} onchange="window.updateTcBranchLabel()"> ${b}</label>`).join('');
            branchDrop.innerHTML = bHtml;
            window.updateTcBranchLabel();
        }
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
            <div class="lc-kpi"><span class="lc-kpi-label">Total OSG Revenue</span><span class="lc-kpi-val">${fmtShortHtml(totalORev)}</span></div>
            <div class="lc-kpi"><span class="lc-kpi-label">Total Prod Revenue</span><span class="lc-kpi-val">${fmtShortHtml(totalPRev)}</span></div>
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
                <td class="number-cell">${fmtShortHtml(e.pRev)}</td>
                <td class="number-cell">${fmtShortHtml(e.oRev)}</td>
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
        const selectedBranches = window.getTcSelectedBranches();
        const selRBM = $('tcRBM').value;
        const selBDM = $('tcBDM').value;
        const allStats = buildStaffStats();
        const filtered = allStats
            .filter(s => s.pQty >= minQty && s[sortBy] > 0)
            .filter(s => !selectedBranches || selectedBranches.includes(s.branch))
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
        const data = filtered.map((e, i) => [
            i + 1, e.name, e.branch, e.rbm, e.bdm, e.pQty, e.oQty,
            parseFloat(e.qtyConv.toFixed(2)), parseFloat(e.valConv.toFixed(2)), Math.round(e.pRev), Math.round(e.oRev)
        ]);
        exportToStyledExcel(data, hdr, 'top_conv_staff.xlsx', 'Top Conversion Staff');
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
        html += insightCard('ÃƒÂ°Ã…Â¸Ã¢â‚¬Å“Ã…Â ', 'Overall Performance Summary', 'info', `
            <div class="insight-metrics">
                <div class="insight-metric"><span class="metric-val">${formatNumber(productData.length)}</span><span class="metric-label">Total Transactions</span></div>
                <div class="insight-metric"><span class="metric-val">${totalStaff}</span><span class="metric-label">Active Staff</span></div>
                <div class="insight-metric"><span class="metric-val">${conv.valueConv.toFixed(2)}%</span><span class="metric-label">Value Conversion</span></div>
                <div class="insight-metric"><span class="metric-val">${conv.qtyConv.toFixed(2)}%</span><span class="metric-label">Qty Conversion</span></div>
                <div class="insight-metric"><span class="metric-val">${fmtShortHtml(totalPRev)}</span><span class="metric-label">Total Prod Revenue</span></div>
                <div class="insight-metric"><span class="metric-val">${Object.keys(pByBranch).length}</span><span class="metric-label">Active Branches</span></div>
            </div>
        `);

        // ---- Card 2: Zero Conversion Alert ----
        if (zeroConvStaff.length > 0) {
            const zeroTotalQty = zeroConvStaff.reduce((s, r) => s + r.pQty, 0);
            const zeroTotalRev = zeroConvStaff.reduce((s, r) => s + r.pRev, 0);
            const topZero = zeroConvStaff.sort((a, b) => b.pQty - a.pQty).slice(0, 5);
            html += insightCard('', `Zero Conversion Alert ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ${zeroConvStaff.length} Staff`, 'danger', `
                <p><strong>${zeroConvStaff.length} staff</strong> have sold <strong>${formatNumber(zeroTotalQty)} products</strong> (${fmtShortHtml(zeroTotalRev)} revenue) but <strong>zero OSG/warranty conversion</strong>.</p>
                <div class="insight-tag-row">
                    ${topZero.map(s => `<span class="insight-tag danger">${s.name} (${s.pQty} qty)</span>`).join('')}
                    ${zeroConvStaff.length > 5 ? `<span class="insight-tag muted">+${zeroConvStaff.length - 5} more</span>` : ''}
                </div>
                <div class="insight-solution">
                    <strong> Solution:</strong> Conduct targeted training for these staff members on OSG selling techniques. Pair them with top converters for mentorship. Set 1-week conversion targets with incentives.
                </div>
            `);
        }

        // ---- Card 3: Top Performers ----
        if (topQty.length > 0) {
            html += insightCard('Ãƒâ€šÃ‚Â ', 'Top Performers ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â Best Qty Conversion', 'success', `
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
                    <strong> Recommendation:</strong> Recognize these staff publicly. Study their techniques and replicate across other branches. Consider a reward/incentive program to sustain performance.
                </div>
            `);
        }

        // ---- Card 4: Underperforming Branches ----
        const weakBranches = branchStats.filter(b => b.pQty >= 10 && b.qtyConv < 2).sort((a, b) => a.qtyConv - b.qtyConv).slice(0, 5);
        if (weakBranches.length > 0) {
            html += insightCard('ÃƒÂ°Ã…Â¸Ã¢â‚¬Å“Ã¢â‚¬Â°', 'Underperforming Branches', 'warning', `
                <p>These branches have significant product sales but very low OSG conversion:</p>
                <table class="data-table insight-table"><thead><tr>
                    <th>Branch</th><th>Prod Qty</th><th>OSG Qty</th><th>Qty Conv%</th>
                </tr></thead><tbody>
                    ${weakBranches.map(b => `<tr><td>${b.name}</td><td class="number-cell">${b.pQty}</td><td class="number-cell">${b.oQty}</td><td class="number-cell loss-val">${b.qtyConv.toFixed(2)}%</td></tr>`).join('')}
                </tbody></table>
                <div class="insight-solution">
                    <strong> Solution:</strong> Schedule branch visits and OSG training workshops. Review branch-level OSG targets. Investigate if product mix or customer demographics contribute to low conversion.
                </div>
            `);
        }

        // ---- Card 5: RBM Performance Gap ----
        if (rbmStats.length >= 2) {
            const rbmSorted = [...rbmStats].sort((a, b) => b.qtyConv - a.qtyConv);
            const best = rbmSorted[0];
            const worst = rbmSorted[rbmSorted.length - 1];
            const gap = best.qtyConv - worst.qtyConv;
            html += insightCard('ÃƒÂ°Ã…Â¸Ã¢â‚¬ËœÃ‚Â¥', 'RBM Performance Gap', gap > 5 ? 'warning' : 'info', `
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
                    <strong> Recommendation:</strong> ${gap > 5 ? 'Organize knowledge-sharing sessions between top and bottom RBMs. Assign mentors and set improvement timelines.' : 'Performance is fairly balanced. Focus on pushing overall numbers higher.'}
                </div>
            `);
        }

        // ---- Card 6: Product Category Analysis ----
        const prodSorted = [...prodStats].filter(p => p.pQty >= 5).sort((a, b) => a.qtyConv - b.qtyConv);
        const weakProds = prodSorted.slice(0, 3);
        const strongProds = prodSorted.slice(-3).reverse();
        if (prodSorted.length > 0) {
            html += insightCard('ÃƒÂ°Ã…Â¸Ã¢â‚¬Å“Ã¢â‚¬Â¹', 'Product Category Analysis', 'info', `
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
                    <strong> Solution:</strong> Focus OSG push on weak categories. Create category-specific sales scripts. Consider bundled OSG offers for low-converting product types.
                </div>
            `);
        }

        // ---- Card 7: Revenue Concentration Risk ----
        if (branchRevShare.length >= 3) {
            const top3Share = branchRevShare.slice(0, 3).reduce((s, b) => s + b.share, 0);
            html += insightCard('ÃƒÂ¢Ã…Â¡Ã¢â‚¬â€œÃƒÂ¯Ã‚Â¸Ã‚Â', 'Revenue Concentration', top3Share > 50 ? 'warning' : 'info', `
                <p>Top 3 branches contribute <strong>${top3Share.toFixed(1)}%</strong> of total product revenue:</p>
                <div class="insight-tag-row">
                    ${branchRevShare.slice(0, 5).map(b => `<span class="insight-tag info">${b.name}: ${b.share.toFixed(1)}%</span>`).join('')}
                </div>
                ${top3Share > 50 ? '<p class="text-warning"> High concentration risk ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â underperformance in these branches would significantly impact overall numbers.</p>' : '<p class="text-success">ÃƒÂ¢Ã…â€œÃ¢â‚¬Â¦ Revenue is fairly distributed ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â good diversification.</p>'}
                <div class="insight-solution">
                    <strong> Recommendation:</strong> ${top3Share > 50 ? 'Invest in growing smaller branches. Reduce dependency on top branches by improving performance of bottom 50%.' : 'Maintain balanced growth across all branches.'}
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

        html += insightCard('ÃƒÂ°Ã…Â¸Ã¢â‚¬ÂÃ‚Â', 'Deep Root Cause Analysis', 'danger', `
            <p style="margin-bottom:1rem; color:var(--text-primary); font-weight:500;">Based on combinatorial data analysis, the primary drivers of lost conversion are:</p>
            ${deepAnalysisHtml}
        `);

        // ---- Card 8: Action Plan ----
        const urgentActions = [];
        if (zeroConvStaff.length > 5) urgentActions.push(`Train ${zeroConvStaff.length} zero-conversion staff on OSG selling immediately`);
        if (weakBranches.length > 0) urgentActions.push(`Conduct branch visits to ${weakBranches.map(b => b.name).join(', ')}`);
        if (conv.qtyConv < 5) urgentActions.push(`Overall qty conversion (${conv.qtyConv.toFixed(1)}%) is below target ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â launch org-wide OSG campaign`);
        urgentActions.push('Review and update staff-wise weekly conversion targets');
        urgentActions.push('Share top performer success stories in team meetings');
        if (topQty.length > 0) urgentActions.push(`Reward top converters: ${topQty.slice(0, 3).map(s => s.name).join(', ')}`);

        html += insightCard('', 'Action Plan ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â Next Steps', 'action', `
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
    $('btnFSExportDashboard').addEventListener('click', () => {
        const modal = document.getElementById('fsExportDashboardModal');
        if (modal) { modal.style.display = 'flex'; renderFsExportDashboard(true); }
    });
    document.querySelector('[data-section="future-section"]').addEventListener('click', () => {
        setTimeout(renderFutureStoresPage, 50);
    });

    function buildFutureStaffStats() {
        const futureProduct = productData.filter(r => r.branch && r.branch.toUpperCase().includes('FUTURE'));

        const pByStaff = {};
        futureProduct.forEach(r => {
            const s = r.staff || 'Unknown';
            if (!pByStaff[s]) pByStaff[s] = { branch: r.branch, rbm: r.rbm, bdm: r.bdm, rows: [] };
            pByStaff[s].rows.push(r);
        });

        // Helper to resolve staff and ensure it's a Future branch
        const processWarr = (r, byStaffObj) => {
            const meta = window.getBranchMeta(r);
            if (!meta && !r.branch) return;
            const branchName = meta ? meta.origBranch : r.branch;
            
            if (!branchName.toUpperCase().includes('FUTURE')) return; // strictly Future stores

            const s = r.staff || ('Unknown|' + branchName);
            if (!byStaffObj[s]) byStaffObj[s] = [];
            byStaffObj[s].push(r);
        };

        const oByStaff = {};
        osgData.forEach(r => processWarr(r, oByStaff));

        const amcByStaff = {};
        amcData.forEach(r => processWarr(r, amcByStaff));

        const samByStaff = {};
        samsungData.forEach(r => processWarr(r, samByStaff));

        const allStaff = new Set([...Object.keys(pByStaff), ...Object.keys(oByStaff), ...Object.keys(amcByStaff), ...Object.keys(samByStaff)]);
        allStaff.delete('Unknown');

        return Array.from(allStaff).map(name => {
            let branch = '';
            let rbm = '';
            let bdm = '';
            let displayName = name;

            if (name.startsWith('Unknown|')) {
                branch = name.split('|')[1] || '';
                displayName = 'Unknown';
                const m = window.getBranchMeta({ branch });
                if (m) {
                    rbm = m.rbm;
                    bdm = m.bdm;
                }
            } else {
                const info = pByStaff[name] || { branch: '', rbm: '', bdm: '' };
                branch = info.branch;
                rbm = info.rbm;
                bdm = info.bdm;
            }

            const pInfo = pByStaff[name] || { branch, rbm, bdm, rows: [] };
            const oRows = oByStaff[name] || [];
            const amcRows = amcByStaff[name] || [];
            const samRows = samByStaff[name] || [];

            const pQty = pInfo.rows.reduce((s, r) => s + r.qty, 0);
            const oQty = oRows.reduce((s, r) => s + r.qty, 0);
            const pRev = pInfo.rows.reduce((s, r) => s + r.soldPrice, 0);
            const oRev = oRows.reduce((s, r) => s + r.soldPrice, 0);
            const qtyConv = pQty > 0 ? (oQty / pQty) * 100 : 0;
            const valConv = pRev > 0 ? (oRev / pRev) * 100 : 0;

            const prodCounts = {};
            const prodRev = {};
            const lgProdCounts = {};
            const samProdCounts = {};
            const lgProdRev = {};
            const samProdRev = {};
            const samsungAllowedCats = ['AC', 'MICROWAVE OVEN', 'REFRIGERATOR', 'WASHING MACHINE', 'TV'];

            pInfo.rows.forEach(r => {
                const originalP = r.product || 'Unknown';
                const p = r.product ? r.product.toUpperCase().trim() : 'UNKNOWN';
                const b = r.brand ? r.brand.toUpperCase() : '';

                prodCounts[originalP] = (prodCounts[originalP] || 0) + (r.qty || 0);
                prodRev[originalP] = (prodRev[originalP] || 0) + (r.soldPrice || 0);

                if (b.includes('LG')) {
                    lgProdCounts[originalP] = (lgProdCounts[originalP] || 0) + (r.qty || 0);
                    lgProdRev[originalP] = (lgProdRev[originalP] || 0) + (r.soldPrice || 0);
                }

                if (b.includes('SAMSUNG') && samsungAllowedCats.includes(p)) {
                    samProdCounts[originalP] = (samProdCounts[originalP] || 0) + (r.qty || 0);
                    samProdRev[originalP] = (samProdRev[originalP] || 0) + (r.soldPrice || 0);
                }
            });

            const oProdCounts = {};
            const oProdRev = {};
            oRows.forEach(r => {
                // OSG category column is already normalized; fall back to normalizing raw product name
                const p = r.category || normalizeProductCategory(r.product) || 'Unknown';
                oProdCounts[p] = (oProdCounts[p] || 0) + (r.qty || 0);
                oProdRev[p] = (oProdRev[p] || 0) + (r.soldPrice || 0);
            });

            const amcProdCounts = {};
            const amcProdRev = {};
            amcRows.forEach(r => {
                const p = r.product || 'Unknown';
                amcProdCounts[p] = (amcProdCounts[p] || 0) + (r.qty || 0);
                amcProdRev[p] = (amcProdRev[p] || 0) + (r.soldPrice || 0);
            });

            const samOsgCounts = {};
            const samOsgRev = {};
            samRows.forEach(r => {
                // Samsung: r.product is already normalized by mapSamsungProductCategory
                const p = r.product || 'Unknown';
                samOsgCounts[p] = (samOsgCounts[p] || 0) + (r.qty || 0);
                samOsgRev[p] = (samOsgRev[p] || 0) + (r.soldPrice || 0);
            });

            const allProds = new Set([...Object.keys(prodCounts), ...Object.keys(oProdCounts), ...Object.keys(amcProdCounts), ...Object.keys(samOsgCounts)]);
            const products = Array.from(allProds).map(p => {
                const q = prodCounts[p] || 0;
                const lgQ = lgProdCounts[p] || 0;
                const samPQ = samProdCounts[p] || 0;

                const oQ = oProdCounts[p] || 0;
                const aQ = amcProdCounts[p] || 0;
                const sQ = samOsgCounts[p] || 0;

                const r = prodRev[p] || 0;
                const lgR = lgProdRev[p] || 0;
                const samPR = samProdRev[p] || 0;

                const oR = oProdRev[p] || 0;
                const aR = amcProdRev[p] || 0;
                const sR = samOsgRev[p] || 0;

                return {
                    name: p,
                    qty: q,
                    lgProdQty: lgQ,
                    samProdQty: samPQ,
                    osgQty: oQ,
                    amcQty: aQ,
                    samQty: sQ,
                    amcRev: aR,
                    samRev: sR,
                    productRev: r,
                    lgProdRev: lgR,
                    samProdRev: samPR,
                    qtyConv: q > 0 ? (oQ / q) * 100 : 0,
                    valConv: r > 0 ? (oR / r) * 100 : 0,
                    amcQtyConv: lgQ > 0 ? (aQ / lgQ) * 100 : 0,
                    amcValConv: lgR > 0 ? (aR / lgR) * 100 : 0,
                    samQtyConv: samPQ > 0 ? (sQ / samPQ) * 100 : 0,
                    samValConv: samPR > 0 ? (sR / samPR) * 100 : 0
                };
            }).sort((a, b) => b.qty - a.qty);

            if (!branch || !rbm || !bdm) {
                const amcFirst = amcRows.find(x => x.branch && x.rbm && x.bdm);
                if (amcFirst) {
                    if (!branch) branch = amcFirst.branch;
                    if (!rbm) rbm = amcFirst.rbm;
                    if (!bdm) bdm = amcFirst.bdm;
                }
            }

            return { name: displayName, branch, rbm, bdm, pQty, oQty, pRev, oRev, qtyConv, valConv, products };
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
        const minOsgQty = parseInt($('lcMinOsgQty').value) || 0;
        const selectedBranches = window.getLcSelectedBranches();
        const selRBM = $('lcRBM').value;
        const selBDM = $('lcBDM').value;

        const allStats = buildStaffStats();

        // Populate dropdowns (preserve selection)
        const branchSet = [...new Set(allStats.map(s => s.branch).filter(Boolean))].sort();
        const rbmSet = [...new Set(allStats.map(s => s.rbm).filter(Boolean))].sort();
        const bdmSet = [...new Set(allStats.map(s => s.bdm).filter(Boolean))].sort();

        const branchDrop = $('lcBranchDropdown');
        const rbmEl = $('lcRBM');
        const bdmEl = $('lcBDM');
        const prevRBM = selRBM;
        const prevBDM = selBDM;

        if (branchDrop) {
            let bHtml = `<label style="display: flex; align-items: center; gap: 8px; padding: 4px 0; cursor: pointer; color: var(--text-primary); text-transform: none; font-weight: 500; font-size: 0.85rem; margin:0;"><input type="checkbox" value="ALL" ${!selectedBranches ? 'checked' : ''} onchange="window.toggleAllLcBranches(this)"> <strong>All Branches</strong></label><hr style="margin: 4px 0; border: none; border-top: 1px solid var(--border);">`;
            bHtml += branchSet.map(b => `<label style="display: flex; align-items: center; gap: 8px; padding: 4px 0; cursor: pointer; color: var(--text-primary); text-transform: none; font-weight: 500; font-size: 0.85rem; margin:0;"><input type="checkbox" value="${b}" class="lc-branch-cb" ${!selectedBranches || selectedBranches.includes(b) ? 'checked' : ''} onchange="window.updateLcBranchLabel()"> ${b}</label>`).join('');
            branchDrop.innerHTML = bHtml;
            window.updateLcBranchLabel();
        }
        rbmEl.innerHTML = '<option value="">All RBMs</option>' +
            rbmSet.map(r => `<option value="${r}" ${r === prevRBM ? 'selected' : ''}>${r}</option>`).join('');
        bdmEl.innerHTML = '<option value="">All BDMs</option>' +
            bdmSet.map(b => `<option value="${b}" ${b === prevBDM ? 'selected' : ''}>${b}</option>`).join('');

        // Filter: minQty, maxConv, minOsgQty, optional Branch, optional RBM, optional BDM
        const filtered = allStats
            .filter(s => s.pQty >= minQty && s.qtyConv <= maxConv)
            .filter(s => s.oQty >= minOsgQty)
            .filter(s => !selectedBranches || selectedBranches.includes(s.branch))
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
            <div class="lc-kpi"><span class="lc-kpi-label">Total Product Revenue</span><span class="lc-kpi-val">${fmtShortHtml(totalPRev)}</span></div>
        `;

        if (filtered.length === 0) {
            $('lcTableWrapper').innerHTML = noDataHTML(`No staff found with â‰¥${minQty} product qty and ÃƒÂ¢Ã¢â‚¬Â°Ã‚Â¤${maxConv}% qty conversion.`);
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
                <td class="number-cell">${fmtShortHtml(e.pRev)}</td>
            </tr>`;
        });

        html += '</tbody></table>';
        $('lcTableWrapper').innerHTML = html;
    }

    function exportLowConvCSV() {
        if (productData.length === 0) return;
        const minQty = parseFloat($('lcMinQty').value) || 0;
        const maxConv = parseFloat($('lcMaxConv').value);
        const minOsgQty = parseInt($('lcMinOsgQty').value) || 0;
        const selectedBranches = window.getLcSelectedBranches();
        const selRBM = $('lcRBM').value;
        const selBDM = $('lcBDM').value;
        const allStats = buildStaffStats();
        const filtered = allStats
            .filter(s => s.pQty >= minQty && s.qtyConv <= maxConv)
            .filter(s => s.oQty >= minOsgQty)
            .filter(s => !selectedBranches || selectedBranches.includes(s.branch))
            .filter(s => !selRBM || s.rbm === selRBM)
            .filter(s => !selBDM || s.bdm === selBDM)
            .sort((a, b) => a.qtyConv - b.qtyConv || b.pQty - a.pQty);
        if (filtered.length === 0) return;
        const hdr = ['Rank', 'Staff', 'Branch', 'RBM', 'BDM', 'Prod Qty', 'OSG Qty', 'Qty Conv%', 'Val Conv%', 'Prod Revenue'];
        const data = filtered.map((e, i) => [
            i + 1, e.name, e.branch, e.rbm, e.bdm, e.pQty, e.oQty,
            parseFloat(e.qtyConv.toFixed(2)), parseFloat(e.valConv.toFixed(2)), Math.round(e.pRev)
        ]);
        exportToStyledExcel(data, hdr, 'low_conv_staff.xlsx', 'Low Conversion Staff');
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
        const selectedBranches = window.getTcSelectedBranches();
        const selRBM = $('tcRBM').value;
        const selBDM = $('tcBDM').value;

        const allStats = buildStaffStats();

        // Populate RBM and BDM dropdowns (preserve selection)
        const branchSet = [...new Set(allStats.map(s => s.branch).filter(Boolean))].sort();
        const rbmSet = [...new Set(allStats.map(s => s.rbm).filter(Boolean))].sort();
        const bdmSet = [...new Set(allStats.map(s => s.bdm).filter(Boolean))].sort();
        const branchDrop = $('tcBranchDropdown');
        if (branchDrop) {
            let bHtml = `<label style="display: flex; align-items: center; gap: 8px; padding: 4px 0; cursor: pointer; color: var(--text-primary); text-transform: none; font-weight: 500; font-size: 0.85rem; margin:0;"><input type="checkbox" value="ALL" ${!selectedBranches ? 'checked' : ''} onchange="window.toggleAllTcBranches(this)"> <strong>All Branches</strong></label><hr style="margin: 4px 0; border: none; border-top: 1px solid var(--border);">`;
            bHtml += branchSet.map(b => `<label style="display: flex; align-items: center; gap: 8px; padding: 4px 0; cursor: pointer; color: var(--text-primary); text-transform: none; font-weight: 500; font-size: 0.85rem; margin:0;"><input type="checkbox" value="${b}" class="tc-branch-cb" ${!selectedBranches || selectedBranches.includes(b) ? 'checked' : ''} onchange="window.updateTcBranchLabel()"> ${b}</label>`).join('');
            branchDrop.innerHTML = bHtml;
            window.updateTcBranchLabel();
        }
        $('tcRBM').innerHTML = '<option value="">All RBMs</option>' +
            rbmSet.map(r => `<option value="${r}" ${r === selRBM ? 'selected' : ''}>${r}</option>`).join('');
        $('tcBDM').innerHTML = '<option value="">All BDMs</option>' +
            bdmSet.map(b => `<option value="${b}" ${b === selBDM ? 'selected' : ''}>${b}</option>`).join('');

        // Filter: must have >= minQty product qty AND conversion > 0, plus Branch/RBM/BDM filters
        const eligible = allStats
            .filter(s => s.pQty >= minQty && s[sortBy] > 0)
            .filter(s => !selectedBranches || selectedBranches.includes(s.branch))
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
            <div class="lc-kpi"><span class="lc-kpi-label">Total OSG Revenue</span><span class="lc-kpi-val">${fmtShortHtml(totalORev)}</span></div>
            <div class="lc-kpi"><span class="lc-kpi-label">Total Prod Revenue</span><span class="lc-kpi-val">${fmtShortHtml(totalPRev)}</span></div>
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
                <td class="number-cell">${fmtShortHtml(e.pRev)}</td>
                <td class="number-cell">${fmtShortHtml(e.oRev)}</td>
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
        const selectedBranches = window.getTcSelectedBranches();
        const selRBM = $('tcRBM').value;
        const selBDM = $('tcBDM').value;
        const allStats = buildStaffStats();
        const filtered = allStats
            .filter(s => s.pQty >= minQty && s[sortBy] > 0)
            .filter(s => !selectedBranches || selectedBranches.includes(s.branch))
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
        const data = filtered.map((e, i) => [
            i + 1, e.name, e.branch, e.rbm, e.bdm, e.pQty, e.oQty,
            parseFloat(e.qtyConv.toFixed(2)), parseFloat(e.valConv.toFixed(2)), Math.round(e.pRev), Math.round(e.oRev)
        ]);
        exportToStyledExcel(data, hdr, 'top_conv_staff.xlsx', 'Top Conversion Staff');
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
        html += insightCard('ÃƒÂ°Ã…Â¸Ã¢â‚¬Å“Ã…Â ', 'Overall Performance Summary', 'info', `
            <div class="insight-metrics">
                <div class="insight-metric"><span class="metric-val">${formatNumber(productData.length)}</span><span class="metric-label">Total Transactions</span></div>
                <div class="insight-metric"><span class="metric-val">${totalStaff}</span><span class="metric-label">Active Staff</span></div>
                <div class="insight-metric"><span class="metric-val">${conv.valueConv.toFixed(2)}%</span><span class="metric-label">Value Conversion</span></div>
                <div class="insight-metric"><span class="metric-val">${conv.qtyConv.toFixed(2)}%</span><span class="metric-label">Qty Conversion</span></div>
                <div class="insight-metric"><span class="metric-val">${fmtShortHtml(totalPRev)}</span><span class="metric-label">Total Prod Revenue</span></div>
                <div class="insight-metric"><span class="metric-val">${Object.keys(pByBranch).length}</span><span class="metric-label">Active Branches</span></div>
            </div>
        `);

        // ---- Card 2: Zero Conversion Alert ----
        if (zeroConvStaff.length > 0) {
            const zeroTotalQty = zeroConvStaff.reduce((s, r) => s + r.pQty, 0);
            const zeroTotalRev = zeroConvStaff.reduce((s, r) => s + r.pRev, 0);
            const topZero = zeroConvStaff.sort((a, b) => b.pQty - a.pQty).slice(0, 5);
            html += insightCard('', `Zero Conversion Alert ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ${zeroConvStaff.length} Staff`, 'danger', `
                <p><strong>${zeroConvStaff.length} staff</strong> have sold <strong>${formatNumber(zeroTotalQty)} products</strong> (${fmtShortHtml(zeroTotalRev)} revenue) but <strong>zero OSG/warranty conversion</strong>.</p>
                <div class="insight-tag-row">
                    ${topZero.map(s => `<span class="insight-tag danger">${s.name} (${s.pQty} qty)</span>`).join('')}
                    ${zeroConvStaff.length > 5 ? `<span class="insight-tag muted">+${zeroConvStaff.length - 5} more</span>` : ''}
                </div>
                <div class="insight-solution">
                    <strong> Solution:</strong> Conduct targeted training for these staff members on OSG selling techniques. Pair them with top converters for mentorship. Set 1-week conversion targets with incentives.
                </div>
            `);
        }

        // ---- Card 3: Top Performers ----
        if (topQty.length > 0) {
            html += insightCard('Ãƒâ€šÃ‚Â ', 'Top Performers ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â Best Qty Conversion', 'success', `
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
                    <strong> Recommendation:</strong> Recognize these staff publicly. Study their techniques and replicate across other branches. Consider a reward/incentive program to sustain performance.
                </div>
            `);
        }

        // ---- Card 4: Underperforming Branches ----
        const weakBranches = branchStats.filter(b => b.pQty >= 10 && b.qtyConv < 2).sort((a, b) => a.qtyConv - b.qtyConv).slice(0, 5);
        if (weakBranches.length > 0) {
            html += insightCard('ÃƒÂ°Ã…Â¸Ã¢â‚¬Å“Ã¢â‚¬Â°', 'Underperforming Branches', 'warning', `
                <p>These branches have significant product sales but very low OSG conversion:</p>
                <table class="data-table insight-table"><thead><tr>
                    <th>Branch</th><th>Prod Qty</th><th>OSG Qty</th><th>Qty Conv%</th>
                </tr></thead><tbody>
                    ${weakBranches.map(b => `<tr><td>${b.name}</td><td class="number-cell">${b.pQty}</td><td class="number-cell">${b.oQty}</td><td class="number-cell loss-val">${b.qtyConv.toFixed(2)}%</td></tr>`).join('')}
                </tbody></table>
                <div class="insight-solution">
                    <strong> Solution:</strong> Schedule branch visits and OSG training workshops. Review branch-level OSG targets. Investigate if product mix or customer demographics contribute to low conversion.
                </div>
            `);
        }

        // ---- Card 5: RBM Performance Gap ----
        if (rbmStats.length >= 2) {
            const rbmSorted = [...rbmStats].sort((a, b) => b.qtyConv - a.qtyConv);
            const best = rbmSorted[0];
            const worst = rbmSorted[rbmSorted.length - 1];
            const gap = best.qtyConv - worst.qtyConv;
            html += insightCard('ÃƒÂ°Ã…Â¸Ã¢â‚¬ËœÃ‚Â¥', 'RBM Performance Gap', gap > 5 ? 'warning' : 'info', `
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
                    <strong> Recommendation:</strong> ${gap > 5 ? 'Organize knowledge-sharing sessions between top and bottom RBMs. Assign mentors and set improvement timelines.' : 'Performance is fairly balanced. Focus on pushing overall numbers higher.'}
                </div>
            `);
        }

        // ---- Card 6: Product Category Analysis ----
        const prodSorted = [...prodStats].filter(p => p.pQty >= 5).sort((a, b) => a.qtyConv - b.qtyConv);
        const weakProds = prodSorted.slice(0, 3);
        const strongProds = prodSorted.slice(-3).reverse();
        if (prodSorted.length > 0) {
            html += insightCard('ÃƒÂ°Ã…Â¸Ã¢â‚¬Å“Ã¢â‚¬Â¹', 'Product Category Analysis', 'info', `
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
                    <strong> Solution:</strong> Focus OSG push on weak categories. Create category-specific sales scripts. Consider bundled OSG offers for low-converting product types.
                </div>
            `);
        }

        // ---- Card 7: Revenue Concentration Risk ----
        if (branchRevShare.length >= 3) {
            const top3Share = branchRevShare.slice(0, 3).reduce((s, b) => s + b.share, 0);
            html += insightCard('ÃƒÂ¢Ã…Â¡Ã¢â‚¬â€œÃƒÂ¯Ã‚Â¸Ã‚Â', 'Revenue Concentration', top3Share > 50 ? 'warning' : 'info', `
                <p>Top 3 branches contribute <strong>${top3Share.toFixed(1)}%</strong> of total product revenue:</p>
                <div class="insight-tag-row">
                    ${branchRevShare.slice(0, 5).map(b => `<span class="insight-tag info">${b.name}: ${b.share.toFixed(1)}%</span>`).join('')}
                </div>
                ${top3Share > 50 ? '<p class="text-warning"> High concentration risk ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â underperformance in these branches would significantly impact overall numbers.</p>' : '<p class="text-success">ÃƒÂ¢Ã…â€œÃ¢â‚¬Â¦ Revenue is fairly distributed ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â good diversification.</p>'}
                <div class="insight-solution">
                    <strong> Recommendation:</strong> ${top3Share > 50 ? 'Invest in growing smaller branches. Reduce dependency on top branches by improving performance of bottom 50%.' : 'Maintain balanced growth across all branches.'}
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

        html += insightCard('ÃƒÂ°Ã…Â¸Ã¢â‚¬ÂÃ‚Â', 'Deep Root Cause Analysis', 'danger', `
            <p style="margin-bottom:1rem; color:var(--text-primary); font-weight:500;">Based on combinatorial data analysis, the primary drivers of lost conversion are:</p>
            ${deepAnalysisHtml}
        `);

        // ---- Card 8: Action Plan ----
        const urgentActions = [];
        if (zeroConvStaff.length > 5) urgentActions.push(`Train ${zeroConvStaff.length} zero-conversion staff on OSG selling immediately`);
        if (weakBranches.length > 0) urgentActions.push(`Conduct branch visits to ${weakBranches.map(b => b.name).join(', ')}`);
        if (conv.qtyConv < 5) urgentActions.push(`Overall qty conversion (${conv.qtyConv.toFixed(1)}%) is below target ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â launch org-wide OSG campaign`);
        urgentActions.push('Review and update staff-wise weekly conversion targets');
        urgentActions.push('Share top performer success stories in team meetings');
        if (topQty.length > 0) urgentActions.push(`Reward top converters: ${topQty.slice(0, 3).map(s => s.name).join(', ')}`);

        html += insightCard('', 'Action Plan ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â Next Steps', 'action', `
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
            <div class="lc-kpi"><span class="lc-kpi-label">Total Product Revenue</span><span class="lc-kpi-val">${fmtShortHtml(totalPRev)}</span></div>
            <div class="lc-kpi"><span class="lc-kpi-label">Total OSG Revenue</span><span class="lc-kpi-val">${fmtShortHtml(totalORev)}</span></div>
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
                <td class="number-cell">${fmtShortHtml(e.pRev)}</td>
                <td class="number-cell">${fmtShortHtml(e.oRev)}</td>
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

    const samAllowed = ['AC', 'MICROWAVE OVEN', 'REFRIGERATOR', 'WASHING MACHINE', 'TV'];
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
                } else {
                    // Content-based detection for OVERVIEW sheet (no hardcoded row indices)
                    const val1Ref = XLSX.utils.encode_cell({r: R, c: 1});
                    const val1 = ws[val1Ref] ? String(ws[val1Ref].v).toUpperCase() : '';
                    if (val0.includes('OVERVIEW')) rType = 'title';
                    else if (val0 === 'PRODUCT' || (val0 === 'RBM' && val1 === 'QTY')) rType = 'header';
                    else if (val0 === 'TOTAL') rType = 'total';
                    else if (val0 !== '') { rType = 'data'; altIdx++; }
                }

                rowTypes[R] = rType;

                for (let C = range.s.c; C <= range.e.c; ++C) {
                    const cellRef = XLSX.utils.encode_cell({r: R, c: C});
                    if (!ws[cellRef]) ws[cellRef] = { t: 's', v: '' };

                    const valStr = String(ws[cellRef].v);
                    valStr.split('\n').forEach(l => { if (l.length + 2 > colWidths[C]) colWidths[C] = l.length + 2; });

                    let cType = rType;

                    let style;
                    if (cType === 'title')  style = JSON.parse(JSON.stringify(sTitle));
                    else if (cType === 'header') style = JSON.parse(JSON.stringify(sHeader));
                    else if (cType === 'total') style = JSON.parse(JSON.stringify(sTotal));
                    else {
                        const idx = groupCol >= 0 ? (groupMap[R] || 0) : altIdx;
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
        const aoa3 = [['FUTURE STORES â€” STORE WISE'], ['BDM', 'Branch', 'Product', 'Product Qty', qtyName, 'Qty Conv%', 'Val Conv%', 'OVERALL Qty Conv%', 'OVERALL Val Conv%']];
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

    buildBrandWorkbook('OSG');
    buildBrandWorkbook('LG_AMC');
    buildBrandWorkbook('SAMSUNG');
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

        // Build invoice lookup from product data ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â only count OSG entries that match a product invoice
        const productInvoices = new Set();
        filtP.forEach(r => { if (r.invoice) productInvoices.add(r.invoice); });

        // Build invoice lookup from ALL OSG data (for missed customers section)
        const osgInvoices = new Set();
        osgData.forEach(r => { if (r.invoice) osgInvoices.add(r.invoice); });
        amcData.forEach(r => { if (r.invoice) osgInvoices.add(r.invoice); });
        samsungData.forEach(r => { if (r.invoice) osgInvoices.add(r.invoice); });

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
            <div class="lc-kpi"><span class="lc-kpi-label">Product Revenue</span><span class="lc-kpi-val">${fmtShortHtml(totalPRev)}</span></div>
            <div class="lc-kpi"><span class="lc-kpi-label">OSG Revenue</span><span class="lc-kpi-val">${fmtShortHtml(totalORev)}</span></div>
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
                    <td class="number-cell">${fmtShortHtml(e.pRev)}</td>
                    <td class="number-cell">${fmtShortHtml(e.oRev)}</td>
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
                    <td class="number-cell">${fmtShortHtml(e.pRev)}</td>
                    <td class="number-cell">${fmtShortHtml(e.oRev)}</td>
                </tr>`;
            });
            html += '</tbody></table>';
            $('pdTopConvTable').innerHTML = html;
        } else {
            $('pdTopConvTable').innerHTML = noDataHTML('No products with enough quantity.');
        }

    }

    // ---- WITHOUT OSG DASHBOARD PAGE ----
    window._wosgActiveTab = 'main';
    window.switchWosgTab = function (tab) {
        window._wosgActiveTab = tab;
        document.querySelectorAll('.wosg-tab').forEach(btn => {
            if (btn.dataset.wosgTab === tab) {
                btn.style.background = 'linear-gradient(135deg,#f97316,#ea580c)';
                btn.style.color = '#fff';
            } else {
                btn.style.background = 'transparent';
                btn.style.color = '#64748b';
            }
        });
        if (tab === 'main') renderWosgMain();
        else if (tab === 'caller') renderWosgDashboard();
        else if (tab === 'daily') renderWosgDaily();
        else if (tab === 'monthly') renderWosgMonthly();
    };

    function renderWosgMain() {
        const container = document.getElementById('wosgReportContainer');
        if (!container) return;
        const allMissed = getMissedInvoices();
        if (allMissed.length === 0) {
            container.innerHTML = '<div style="text-align:center;padding:60px;color:var(--text-muted);">Upload data and generate reports first.</div>';
            return;
        }

        const data = {
            date: 'OVERALL',
            sKey: 'OVERALL',
            rows: [],
            totalCount: 0,
            totalValue: 0
        };

        allMissed.forEach(r => {
            const val = Math.abs(r.soldPrice || 0);
            const st = coStatusMap[fbKey(r.invoice)] || {};
            const cs = st.callStatus || '';
            const interest = st.interest || '';

            data.totalCount++;
            data.totalValue += val;

            let statusLabel = '';
            if (interest === 'interested') statusLabel = 'Interested';
            else if (interest === 'not-interested') statusLabel = 'Not Interested';
            else if (interest === 'follow-up') statusLabel = 'Follow-up';
            else if (interest === 'bought') statusLabel = 'Closed';
            else if (cs === 'connected') statusLabel = 'Connected';
            else if (cs === 'disconnected') statusLabel = 'Disconnected';
            else if (cs === 'not-connected') statusLabel = 'Not Connected';
            else statusLabel = 'Not Called';

            const existingRow = data.rows.find(x => x.status === statusLabel);
            if (existingRow) {
                existingRow.count++;
                existingRow.value += val;
            } else {
                data.rows.push({ status: statusLabel, count: 1, value: val });
            }
        });

        window._wosgMainData = [data]; // Stored as array to reuse generic builder
        buildGenericWosgTable(container, [data], 'SUMMARY', 'WITHOUT OSG MAIN REPORT');
    }

    document.querySelector('[data-section="wosg-dashboard-section"]').addEventListener('click', () => {
        loadCoStatuses(() => setTimeout(() => {
            if (window._wosgActiveTab === 'main') renderWosgMain();
            else if (window._wosgActiveTab === 'caller') renderWosgDashboard();
            else if (window._wosgActiveTab === 'daily') renderWosgDaily();
            else if (window._wosgActiveTab === 'monthly') renderWosgMonthly();
        }, 50));
    });



    function getMissedInvoices() {
        if (window.sharedMissedUnique && window.sharedMissedUnique.length > 0) {
            return window.sharedMissedUnique;
        }
        if (productData.length === 0) return [];
        const osgInv = new Set();
        osgData.forEach(r => { if (r.invoice) osgInv.add(r.invoice); });
        amcData.forEach(r => { if (r.invoice) osgInv.add(r.invoice); });
        samsungData.forEach(r => { if (r.invoice) osgInv.add(r.invoice); });

        const seenInv = new Set();
        const allMissed = [];
        productData.forEach(r => {
            if (r.invoice && !osgInv.has(r.invoice) && !seenInv.has(r.invoice)) {
                seenInv.add(r.invoice);
                allMissed.push(r);
            }
        });
        return allMissed;
    }

    function renderWosgDaily() {
        const container = document.getElementById('wosgReportContainer');
        if (!container) return;
        const allMissed = getMissedInvoices();
        if (allMissed.length === 0) {
            container.innerHTML = '<div style="text-align:center;padding:60px;color:var(--text-muted);">Upload data and generate reports first.</div>';
            return;
        }

        // Group by Date
        const dateData = {};
        allMissed.forEach(r => {
            let dtStr = '';
            let dt = r.invoiceDate || r.time;
            if (dt) {
                if (dt instanceof Date && !isNaN(dt)) {
                    dtStr = String(dt.getDate()).padStart(2, '0') + '-' + String(dt.getMonth() + 1).padStart(2, '0') + '-' + dt.getFullYear();
                } else if (typeof dt === 'string') {
                    if (dt.match(/^\d{4}-\d{2}-\d{2}/)) {
                        const parts = dt.substring(0, 10).split('-');
                        dtStr = parts[2] + '-' + parts[1] + '-' + parts[0];
                    } else {
                        dtStr = dt; // fallback
                    }
                } else if (typeof dt === 'number') {
                    let dObj = new Date(Math.round((dt - 25569) * 86400 * 1000));
                    dtStr = String(dObj.getUTCDate()).padStart(2, '0') + '-' + String(dObj.getUTCMonth() + 1).padStart(2, '0') + '-' + dObj.getUTCFullYear();
                }
            }
            if (!dtStr) dtStr = 'Unknown Date';

            if (!dateData[dtStr]) {
                dateData[dtStr] = {
                    date: dtStr,
                    rows: [],
                    totalCount: 0,
                    totalValue: 0,
                    timestamp: (dt instanceof Date && !isNaN(dt)) ? dt.getTime() : 0
                };
            }

            const cd = dateData[dtStr];
            const val = Math.abs(r.soldPrice || 0);

            const st = coStatusMap[fbKey(r.invoice)] || {};
            const cs = st.callStatus || '';
            const interest = st.interest || '';

            cd.totalCount++;
            cd.totalValue += val;

            let statusLabel = '';
            if (interest === 'interested') statusLabel = 'Interested';
            else if (interest === 'not-interested') statusLabel = 'Not Interested';
            else if (interest === 'follow-up') statusLabel = 'Follow-up';
            else if (interest === 'bought') statusLabel = 'Closed';
            else if (cs === 'connected') statusLabel = 'Connected';
            else if (cs === 'disconnected') statusLabel = 'Disconnected';
            else if (cs === 'not-connected') statusLabel = 'Not Connected';
            else statusLabel = 'Not Called';

            const existingRow = cd.rows.find(x => x.status === statusLabel);
            if (existingRow) {
                existingRow.count++;
                existingRow.value += val;
            } else {
                cd.rows.push({ status: statusLabel, count: 1, value: val });
            }
        });

        // Sort dates descending
        const sortedDates = Object.values(dateData).sort((a, b) => {
            if (a.date === 'Unknown Date') return 1;
            if (b.date === 'Unknown Date') return -1;
            // Parse DD-MM-YYYY
            const ap = a.date.split('-');
            const bp = b.date.split('-');
            if (ap.length === 3 && bp.length === 3) {
                const da = new Date(ap[2], ap[1] - 1, ap[0]);
                const db = new Date(bp[2], bp[1] - 1, bp[0]);
                return db.getTime() - da.getTime();
            }
            return 0;
        });

        // Store for Export
        window._wosgDailyData = sortedDates;

        buildGenericWosgTable(container, sortedDates, 'DATE', 'WITHOUT OSG DAILY REPORT');
    }

    function renderWosgMonthly() {
        const container = document.getElementById('wosgReportContainer');
        if (!container) return;
        const allMissed = getMissedInvoices();
        if (allMissed.length === 0) {
            container.innerHTML = '<div style="text-align:center;padding:60px;color:var(--text-muted);">Upload data and generate reports first.</div>';
            return;
        }

        const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

        const monthData = {};
        allMissed.forEach(r => {
            let mStr = '';
            let sKey = '';
            let dt = r.invoiceDate || r.time;
            if (dt) {
                if (dt instanceof Date && !isNaN(dt)) {
                    mStr = monthNames[dt.getMonth()] + ' ' + dt.getFullYear();
                    sKey = dt.getFullYear() + String(dt.getMonth()).padStart(2, '0');
                } else if (typeof dt === 'string') {
                    if (dt.match(/^\d{4}-\d{2}-\d{2}/)) {
                        const parts = dt.substring(0, 10).split('-');
                        mStr = monthNames[parseInt(parts[1]) - 1] + ' ' + parts[0];
                        sKey = parts[0] + parts[1];
                    }
                } else if (typeof dt === 'number') {
                    let dObj = new Date(Math.round((dt - 25569) * 86400 * 1000));
                    mStr = monthNames[dObj.getUTCMonth()] + ' ' + dObj.getUTCFullYear();
                    sKey = dObj.getUTCFullYear() + String(dObj.getUTCMonth()).padStart(2, '0');
                }
            }
            if (!mStr) { mStr = 'Unknown Month'; sKey = '000000'; }

            if (!monthData[mStr]) {
                monthData[mStr] = {
                    date: mStr,
                    sKey: sKey,
                    rows: [],
                    totalCount: 0,
                    totalValue: 0
                };
            }

            const cd = monthData[mStr];
            const val = Math.abs(r.soldPrice || 0);

            const st = coStatusMap[fbKey(r.invoice)] || {};
            const cs = st.callStatus || '';
            const interest = st.interest || '';

            cd.totalCount++;
            cd.totalValue += val;

            let statusLabel = '';
            if (interest === 'interested') statusLabel = 'Interested';
            else if (interest === 'not-interested') statusLabel = 'Not Interested';
            else if (interest === 'follow-up') statusLabel = 'Follow-up';
            else if (interest === 'bought') statusLabel = 'Closed';
            else if (cs === 'connected') statusLabel = 'Connected';
            else if (cs === 'disconnected') statusLabel = 'Disconnected';
            else if (cs === 'not-connected') statusLabel = 'Not Connected';
            else statusLabel = 'Not Called';

            const existingRow = cd.rows.find(x => x.status === statusLabel);
            if (existingRow) {
                existingRow.count++;
                existingRow.value += val;
            } else {
                cd.rows.push({ status: statusLabel, count: 1, value: val });
            }
        });

        const sortedMonths = Object.values(monthData).sort((a, b) => b.sKey.localeCompare(a.sKey));

        // Store for Export
        window._wosgMonthlyData = sortedMonths;

        buildGenericWosgTable(container, sortedMonths, 'MONTH', 'WITHOUT OSG MONTHLY REPORT');
    }

    function renderWosgDashboard() {
        const container = document.getElementById('wosgReportContainer');
        if (!container) return;
        const allMissed = getMissedInvoices();
        if (allMissed.length === 0) {
            container.innerHTML = '<div style="text-align:center;padding:60px;color:var(--text-muted);">Upload data and generate reports first.</div>';
            return;
        }

        const today = new Date();
        const dateStr = String(today.getDate()).padStart(2, '0') + '-' + String(today.getMonth() + 1).padStart(2, '0') + '-' + today.getFullYear();
        const selectedCaller = window._wosgCallerFilter || '';

        const allCallers = Array.from(new Set(CO_CALLERS.map(c => c.name))).sort();
        const callers = selectedCaller ? [selectedCaller] : allCallers;

        const callerData = {};
        callers.forEach(c => callerData[c] = { rows: [], totalCount: 0, totalValue: 0 });

        allMissed.forEach(r => {
            const st = coStatusMap[fbKey(r.invoice)] || {};
            const callerName = st.calledBy || '';

            if (!callerName) return; // skip unassigned
            if (selectedCaller && callerName !== selectedCaller) return;
            if (!callerData[callerName]) callerData[callerName] = { rows: [], totalCount: 0, totalValue: 0 };

            const val = Math.abs(r.soldPrice || 0);
            const cs = st.callStatus || '';
            const interest = st.interest || '';

            callerData[callerName].totalCount++;
            callerData[callerName].totalValue += val;

            let statusLabel = '';
            if (interest === 'interested') statusLabel = 'Interested';
            else if (interest === 'not-interested') statusLabel = 'Not Interested';
            else if (interest === 'follow-up') statusLabel = 'Follow-up';
            else if (interest === 'bought') statusLabel = 'Closed';
            else if (cs === 'connected') statusLabel = 'Connected';
            else if (cs === 'disconnected') statusLabel = 'Disconnected';
            else if (cs === 'not-connected') statusLabel = 'Not Connected';
            else statusLabel = 'Not Called';

            const existingRow = callerData[callerName].rows.find(x => x.status === statusLabel);
            if (existingRow) {
                existingRow.count++;
                existingRow.value += val;
            } else {
                callerData[callerName].rows.push({ status: statusLabel, count: 1, value: val });
            }
        });

        // Store filtered data for export
        window._wosgCallerData = callerData;
        window._wosgCallers = callers;
        window._wosgCallerFilter_active = selectedCaller;

        // Expose global handler so the inline onchange can call back in
        window.filterCallerReport = function (val) {
            window._wosgCallerFilter = val;
            renderWosgDashboard();
        };

        // Build professional pivoted table
        const cellBorder = '1px solid #334155';
        let grandCount = 0, grandValue = 0;
        const cols = ['Connected', 'Disconnected', 'Not Connected', 'Interested', 'Not Interested', 'Follow-up', 'Closed'];

        // Caller filter dropdown bar
        const callerOptions = allCallers.map(n =>
            `<option value="${n}" ${selectedCaller === n ? 'selected' : ''}>${n}</option>`
        ).join('');
        const filterBar = `
        <div style="display:flex;align-items:center;gap:12px;margin-bottom:16px;background:var(--bg-card);border:1px solid #334155;border-radius:10px;padding:12px 16px;flex-wrap:wrap;">
            <span style="font-size:0.8rem;font-weight:700;color:#f97316;text-transform:uppercase;letter-spacing:1px;">&#128100; Filter by Caller</span>
            <select id="callerReportDropdown" onchange="window.filterCallerReport(this.value)"
                style="background:#0f172a;color:#e2e8f0;border:1px solid #f97316;border-radius:6px;padding:8px 14px;font-size:0.88rem;font-weight:600;font-family:inherit;cursor:pointer;outline:none;min-width:160px;">
                <option value="" ${!selectedCaller ? 'selected' : ''}>All Callers</option>
                ${callerOptions}
            </select>
            ${selectedCaller ? `<span style="background:rgba(249,115,22,0.15);border:1px solid #f97316;color:#f97316;border-radius:6px;padding:4px 10px;font-size:0.8rem;font-weight:700;">Showing: ${selectedCaller}</span>
            <button onclick="window.filterCallerReport('')" style="background:rgba(239,68,68,0.1);border:1px solid rgba(239,68,68,0.3);color:#ef4444;border-radius:6px;padding:6px 12px;font-size:0.8rem;font-weight:700;cursor:pointer;font-family:inherit;">&#10005; Clear Filter</button>` : ''}
        </div>`;

        let html = `
        <div style="overflow-x:auto; border-radius:12px; border:2px solid #334155; box-shadow:0 4px 24px rgba(0,0,0,0.3);">
        <table id="wosgReportTable" style="width:100%;border-collapse:collapse;font-family:'Inter',sans-serif;font-size:0.88rem;background:var(--bg-card);text-align:center;">
        <thead>
        <tr><td colspan="${cols.length + 3}" style="
            background:linear-gradient(135deg,#f59e0b,#ea580c);
            color:#fff; font-size:1.1rem; font-weight:800;
            text-align:center; padding:16px 10px;
            letter-spacing:1px; text-transform:uppercase;
            border-bottom:3px solid #c2410c;
        ">WITHOUT OSG CALLER REPORT &mdash; ${dateStr}${selectedCaller ? ' \u2014 ' + selectedCaller : ''}</td></tr>
        <tr style="background:#0f172a;">
            <th style="padding:12px 14px;text-align:left;color:#f97316;font-weight:700;font-size:0.75rem;letter-spacing:1px;text-transform:uppercase;border:${cellBorder};width:140px;">CALLER</th>
            ${cols.map(c => `<th style="padding:12px 6px;color:#f97316;font-weight:700;font-size:0.75rem;letter-spacing:1px;text-transform:uppercase;border:${cellBorder};">${c}</th>`).join('')}
            <th style="padding:12px 10px;text-align:right;color:#f97316;font-weight:700;font-size:0.75rem;letter-spacing:1px;text-transform:uppercase;border:${cellBorder};">TOTAL VAL</th>
            <th style="padding:12px 10px;color:#f97316;font-weight:700;font-size:0.75rem;letter-spacing:1px;text-transform:uppercase;border:${cellBorder};">TOTAL CNT</th>
        </tr>
        </thead>
        <tbody>`;

        let colTotals = {};
        cols.forEach(c => colTotals[c] = 0);

        callers.forEach((callerName, ci) => {
            const cd = callerData[callerName];
            const callerCfg = CO_CALLERS.find(c => c.name === callerName) || { color: '#f97316', bg: 'rgba(249,115,22,0.15)' };
            const rowBg = ci % 2 === 0 ? '#0f172a' : '#1e293b';

            let counts = {};
            cd.rows.forEach(r => counts[r.status] = r.count);

            grandCount += cd.totalCount;
            grandValue += cd.totalValue;

            html += `<tr style="background:${rowBg}; transition:background 0.2s;" onmouseover="this.style.background='rgba(249,115,22,0.1)';" onmouseout="this.style.background='${rowBg}';">`;
            html += `<td style="padding:12px 14px;border:${cellBorder};vertical-align:middle;text-align:left;background:${callerCfg.bg};border-left:4px solid ${callerCfg.color};">
                <div style="display:flex;align-items:center;gap:10px;">
                    <div style="width:28px;height:28px;border-radius:50%;background:${callerCfg.color};color:#fff;display:flex;align-items:center;justify-content:center;font-size:0.9rem;font-weight:800;flex-shrink:0;">${callerName[0]}</div>
                    <div style="font-weight:700;font-size:0.85rem;color:${callerCfg.color};">${callerName}</div>
                </div>
            </td>`;

            cols.forEach(c => {
                const cnt = counts[c] || 0;
                colTotals[c] += cnt;
                html += `<td style="padding:12px 6px;border:${cellBorder};font-weight:${cnt > 0 ? '700' : '500'};color:${cnt > 0 ? '#e2e8f0' : '#475569'};font-size:0.9rem;">${cnt > 0 ? cnt : '-'}</td>`;
            });

            html += `<td style="padding:12px 10px;border:${cellBorder};text-align:right;font-weight:800;color:${callerCfg.color};font-family:'JetBrains Mono',monospace;font-size:0.9rem;">${fmtShortHtml(cd.totalValue)}</td>`;
            html += `<td style="padding:12px 10px;border:${cellBorder};font-weight:800;color:${callerCfg.color};font-size:0.95rem;">${cd.totalCount}</td>`;
            html += `</tr>`;
        });

        html += `<tr style="background:linear-gradient(135deg,#ea580c,#dc2626);">
            <td style="padding:14px 14px;font-weight:800;color:#fff;font-size:0.85rem;letter-spacing:1px;text-transform:uppercase;border:1px solid rgba(255,255,255,0.15);text-align:left;">GRAND TOTAL</td>`;

        cols.forEach(c => {
            html += `<td style="padding:14px 6px;font-weight:800;color:#fff;font-size:0.95rem;border:1px solid rgba(255,255,255,0.15);">${colTotals[c] > 0 ? colTotals[c] : '-'}</td>`;
        });

        html += `<td style="padding:14px 10px;text-align:right;font-weight:800;color:#fff;font-size:1rem;font-family:'JetBrains Mono',monospace;border:1px solid rgba(255,255,255,0.15);">${fmtShortHtml(grandValue)}</td>
            <td style="padding:14px 10px;font-weight:800;color:#fff;font-size:1.1rem;border:1px solid rgba(255,255,255,0.15);">${grandCount}</td>
            </tr>`;

        html += '</tbody></table></div>';

        const kpiHtml = `
        <div style="display:flex;gap:14px;flex-wrap:wrap;margin-bottom:24px;">
            <div style="flex:1;min-width:140px;background:var(--bg-card);border:1px solid var(--border);border-radius:12px;padding:16px 20px;border-top:3px solid #f97316;">
                <div style="font-size:1.8rem;font-weight:800;color:#f97316;">${grandCount.toLocaleString()}</div>
                <div style="font-size:0.78rem;color:var(--text-muted);margin-top:3px;">Total Records${selectedCaller ? ' \u2014 ' + selectedCaller : ''}</div>
            </div>
            <div style="flex:1;min-width:140px;background:var(--bg-card);border:1px solid var(--border);border-radius:12px;padding:16px 20px;border-top:3px solid #16a34a;">
                <div style="font-size:1.8rem;font-weight:800;color:#16a34a;">${fmtShortHtml(grandValue)}</div>
                <div style="font-size:0.78rem;color:var(--text-muted);margin-top:3px;">Total Value</div>
            </div>
        </div>`;

        container.innerHTML = filterBar + kpiHtml + html;
    }

    function buildGenericWosgTable(container, dataArr, groupColName, title) {
        const today = new Date();
        const dateStr = String(today.getDate()).padStart(2, '0') + '-' + String(today.getMonth() + 1).padStart(2, '0') + '-' + today.getFullYear();

        const cols = ['Connected', 'Disconnected', 'Not Connected', 'Interested', 'Not Interested', 'Follow-up', 'Closed', 'Not Called'];

        const cellBorder = '1px solid #334155';
        let grandCount = 0, grandValue = 0;
        let colTotals = {};
        cols.forEach(c => colTotals[c] = 0);

        let html = `
        <div style="overflow-x:auto; border-radius:12px; border:2px solid #334155; box-shadow:0 4px 24px rgba(0,0,0,0.3);">
        <table id="wosgReportTable" style="width:100%;border-collapse:collapse;font-family:'Inter',sans-serif;font-size:0.88rem;background:var(--bg-card);text-align:center;">
        <thead>
        <tr><td colspan="${cols.length + 3}" style="
            background:linear-gradient(135deg,#2563eb,#1d4ed8);
            color:#fff; font-size:1.1rem; font-weight:800;
            text-align:center; padding:16px 10px;
            letter-spacing:1px; text-transform:uppercase;
            border-bottom:3px solid #1e3a8a;
        ">${title} &mdash; ${dateStr}</td></tr>
        <tr style="background:#0f172a;">
            <th style="padding:12px 14px;text-align:left;color:#60a5fa;font-weight:700;font-size:0.75rem;letter-spacing:1px;text-transform:uppercase;border:${cellBorder};width:150px;">${groupColName}</th>
            ${cols.map(c => `<th style="padding:12px 6px;color:#60a5fa;font-weight:700;font-size:0.75rem;letter-spacing:1px;text-transform:uppercase;border:${cellBorder};">${c}</th>`).join('')}
            <th style="padding:12px 10px;text-align:right;color:#60a5fa;font-weight:700;font-size:0.75rem;letter-spacing:1px;text-transform:uppercase;border:${cellBorder};">TOTAL VAL</th>
            <th style="padding:12px 10px;color:#60a5fa;font-weight:700;font-size:0.75rem;letter-spacing:1px;text-transform:uppercase;border:${cellBorder};">TOTAL CNT</th>
        </tr>
        </thead>
        <tbody>`;

        dataArr.forEach((grp, ci) => {
            const rowBg = ci % 2 === 0 ? '#0f172a' : '#1e293b';
            let counts = {};
            grp.rows.forEach(r => counts[r.status] = r.count);

            grandCount += grp.totalCount;
            grandValue += grp.totalValue;

            html += `<tr style="background:${rowBg}; transition:background 0.2s;" onmouseover="this.style.background='rgba(37,99,235,0.1)';" onmouseout="this.style.background='${rowBg}';">`;

            html += `<td style="padding:12px 14px;border:${cellBorder};vertical-align:middle;text-align:left;background:rgba(37,99,235,0.05);border-left:4px solid #3b82f6;">
                <div style="font-weight:800;font-size:0.9rem;color:#60a5fa;">${grp.date === 'OVERALL' ? 'Company Total' : grp.date}</div>
            </td>`;

            cols.forEach(c => {
                const cnt = counts[c] || 0;
                colTotals[c] += cnt;
                html += `<td style="padding:12px 6px;border:${cellBorder};font-weight:${cnt > 0 ? '700' : '500'};color:${cnt > 0 ? '#e2e8f0' : '#475569'};font-size:0.9rem;">${cnt > 0 ? cnt : '-'}</td>`;
            });

            html += `<td style="padding:12px 10px;border:${cellBorder};text-align:right;font-weight:800;color:#60a5fa;font-family:'JetBrains Mono',monospace;font-size:0.9rem;">${fmtShortHtml(grp.totalValue)}</td>`;
            html += `<td style="padding:12px 10px;border:${cellBorder};font-weight:800;color:#60a5fa;font-size:0.95rem;">${grp.totalCount}</td>`;
            html += '</tr>';
        });

        html += `<tr style="background:linear-gradient(135deg,#1d4ed8,#1e3a8a);">
            <td style="padding:14px 14px;font-weight:800;color:#fff;font-size:0.85rem;letter-spacing:1px;text-transform:uppercase;border:1px solid rgba(255,255,255,0.15);text-align:left;">GRAND TOTAL</td>`;

        cols.forEach(c => {
            html += `<td style="padding:14px 6px;font-weight:800;color:#fff;font-size:0.95rem;border:1px solid rgba(255,255,255,0.15);">${colTotals[c] > 0 ? colTotals[c] : '-'}</td>`;
        });

        html += `<td style="padding:14px 10px;text-align:right;font-weight:800;color:#fff;font-size:1rem;font-family:'JetBrains Mono',monospace;border:1px solid rgba(255,255,255,0.15);">${fmtShortHtml(grandValue)}</td>
            <td style="padding:14px 10px;font-weight:800;color:#fff;font-size:1.1rem;border:1px solid rgba(255,255,255,0.15);">${grandCount}</td>
        </tr>`;

        html += '</tbody></table></div>';

        const kpiHtml = `
        <div style="display:flex;gap:14px;flex-wrap:wrap;margin-bottom:24px;">
            <div style="flex:1;min-width:140px;background:var(--bg-card);border:1px solid var(--border);border-radius:12px;padding:16px 20px;border-top:3px solid #2563eb;">
                <div style="font-size:1.8rem;font-weight:800;color:#3b82f6;">${grandCount.toLocaleString()}</div>
                <div style="font-size:0.78rem;color:var(--text-muted);margin-top:3px;">Total Records</div>
            </div>
            <div style="flex:1;min-width:140px;background:var(--bg-card);border:1px solid var(--border);border-radius:12px;padding:16px 20px;border-top:3px solid #16a34a;">
                <div style="font-size:1.8rem;font-weight:800;color:#16a34a;">${fmtShortHtml(grandValue)}</div>
                <div style="font-size:0.78rem;color:var(--text-muted);margin-top:3px;">Total Value</div>
            </div>
        </div>`;

        container.innerHTML = kpiHtml + html;
    }


    window.exportWosgReport = function () {
        const tab = window._wosgActiveTab || 'main';
        const today = new Date();
        const dateStr = String(today.getDate()).padStart(2, '0') + '-' + String(today.getMonth() + 1).padStart(2, '0') + '-' + today.getFullYear();

        // ---- CALLER TAB: Full styled export ----
        if (tab === 'caller') {
            if (!window._wosgCallerData || !window._wosgCallers) return;

            const selectedCaller = window._wosgCallerFilter_active || '';
            const callers = window._wosgCallers;
            const callerData = window._wosgCallerData;
            const cols = ['Connected', 'Disconnected', 'Not Connected', 'Interested', 'Not Interested', 'Follow-up', 'Closed'];
            const fileName = selectedCaller
                ? `Caller_Report_${selectedCaller}_${dateStr}.xlsx`
                : `Caller_Report_All_${dateStr}.xlsx`;
            const sheetTitle = selectedCaller
                ? `WITHOUT OSG CALLER REPORT â€”Â ${dateStr} â€”Â ${selectedCaller.toUpperCase()}`
                : `WITHOUT OSG CALLER REPORT â€”Â ${dateStr}`;

            // Build AOA
            const aoa = [];
            aoa.push([sheetTitle, ...Array(cols.length + 2).fill('')]);
            aoa.push(['CALLER', ...cols, 'TOTAL VAL', 'TOTAL CNT']);

            let grandCount = 0, grandValue = 0;
            const colTotals = {};
            cols.forEach(c => colTotals[c] = 0);

            callers.forEach(callerName => {
                const cd = callerData[callerName];
                if (!cd) return;
                let counts = {};
                cd.rows.forEach(r => counts[r.status] = r.count);
                let row = [callerName];
                cols.forEach(c => {
                    const v = counts[c] || 0;
                    colTotals[c] += v;
                    row.push(v > 0 ? v : '-');
                });
                row.push(Math.round(cd.totalValue));
                row.push(cd.totalCount);
                aoa.push(row);
                grandCount += cd.totalCount;
                grandValue += cd.totalValue;
            });

            // Grand Total row
            let gtRow = ['GRAND TOTAL'];
            cols.forEach(c => gtRow.push(colTotals[c] > 0 ? colTotals[c] : '-'));
            gtRow.push(Math.round(grandValue));
            gtRow.push(grandCount);
            aoa.push(gtRow);

            const wb = XLSX.utils.book_new();
            const ws = XLSX.utils.aoa_to_sheet(aoa);

            // Column widths
            ws['!cols'] = [{ wch: 20 }, ...cols.map(() => ({ wch: 15 })), { wch: 16 }, { wch: 12 }];
            // Merge title row
            ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: cols.length + 2 } }];
            // Row heights
            ws['!rows'] = [{ hpt: 30 }, { hpt: 24 }, ...callers.map(() => ({ hpt: 22 })), { hpt: 24 }];

            // ---- STYLES ----
            const titleStyle = {
                fill: { patternType: 'solid', fgColor: { rgb: 'EA580C' } },
                font: { color: { rgb: 'FFFFFF' }, bold: true, sz: 13, name: 'Calibri' },
                alignment: { horizontal: 'center', vertical: 'center' },
                border: { bottom: { style: 'medium', color: { rgb: 'C2410C' } } }
            };
            const headerStyle = {
                fill: { patternType: 'solid', fgColor: { rgb: '0F172A' } },
                font: { color: { rgb: 'F97316' }, bold: true, sz: 10, name: 'Calibri' },
                alignment: { horizontal: 'center', vertical: 'center' },
                border: {
                    top: { style: 'thin', color: { rgb: '334155' } },
                    bottom: { style: 'thin', color: { rgb: '334155' } },
                    left: { style: 'thin', color: { rgb: '334155' } },
                    right: { style: 'thin', color: { rgb: '334155' } }
                }
            };
            const grandTotalStyle = {
                fill: { patternType: 'solid', fgColor: { rgb: 'DC2626' } },
                font: { color: { rgb: 'FFFFFF' }, bold: true, sz: 11, name: 'Calibri' },
                alignment: { horizontal: 'center', vertical: 'center' },
                border: {
                    top: { style: 'thin', color: { rgb: 'FFFFFF' } },
                    bottom: { style: 'thin', color: { rgb: 'FFFFFF' } },
                    left: { style: 'thin', color: { rgb: 'FFFFFF' } },
                    right: { style: 'thin', color: { rgb: 'FFFFFF' } }
                }
            };

            // Caller color map from CO_CALLERS config
            const callerColors = {};
            if (typeof CO_CALLERS !== 'undefined') {
                CO_CALLERS.forEach(c => { callerColors[c.name] = c.color || '#F97316'; });
            }
            const hexToRGB = (hex) => hex.replace('#', '').toUpperCase().padStart(6, '0');

            const totalRows = aoa.length;
            const totalCols = cols.length + 3;

            for (let R = 0; R < totalRows; R++) {
                for (let C = 0; C < totalCols; C++) {
                    const addr = XLSX.utils.encode_cell({ r: R, c: C });
                    if (!ws[addr]) ws[addr] = { t: 's', v: '' };

                    if (R === 0) {
                        // Title row
                        ws[addr].s = titleStyle;
                    } else if (R === 1) {
                        // Header row
                        ws[addr].s = headerStyle;
                    } else if (R === totalRows - 1) {
                        // Grand Total row
                        ws[addr].s = grandTotalStyle;
                        if (C === 0) {
                            ws[addr].s = { ...grandTotalStyle, alignment: { horizontal: 'left', vertical: 'center' } };
                        }
                    } else {
                        // Data rows
                        const callerIdx = R - 2;
                        const callerName = callers[callerIdx] || '';
                        const callerHex = hexToRGB(callerColors[callerName] || '#F97316');
                        const rowBg = callerIdx % 2 === 0 ? '0F172A' : '1E293B';

                        if (C === 0) {
                            // Caller name cell â€”Â colored left border + avatar-style bg
                            ws[addr].s = {
                                fill: { patternType: 'solid', fgColor: { rgb: rowBg } },
                                font: { color: { rgb: callerHex }, bold: true, sz: 10, name: 'Calibri' },
                                alignment: { horizontal: 'left', vertical: 'center' },
                                border: {
                                    left: { style: 'medium', color: { rgb: callerHex } },
                                    right: { style: 'thin', color: { rgb: '334155' } },
                                    top: { style: 'thin', color: { rgb: '334155' } },
                                    bottom: { style: 'thin', color: { rgb: '334155' } }
                                }
                            };
                        } else if (C === totalCols - 2) {
                            // Total Value cell
                            ws[addr].s = {
                                fill: { patternType: 'solid', fgColor: { rgb: rowBg } },
                                font: { color: { rgb: callerHex }, bold: true, sz: 10, name: 'Calibri' },
                                alignment: { horizontal: 'right', vertical: 'center' },
                                border: { top: { style: 'thin', color: { rgb: '334155' } }, bottom: { style: 'thin', color: { rgb: '334155' } }, left: { style: 'thin', color: { rgb: '334155' } }, right: { style: 'thin', color: { rgb: '334155' } } }
                            };
                        } else if (C === totalCols - 1) {
                            // Total Count cell
                            ws[addr].s = {
                                fill: { patternType: 'solid', fgColor: { rgb: rowBg } },
                                font: { color: { rgb: callerHex }, bold: true, sz: 11, name: 'Calibri' },
                                alignment: { horizontal: 'center', vertical: 'center' },
                                border: { top: { style: 'thin', color: { rgb: '334155' } }, bottom: { style: 'thin', color: { rgb: '334155' } }, left: { style: 'thin', color: { rgb: '334155' } }, right: { style: 'thin', color: { rgb: '334155' } } }
                            };
                        } else {
                            // Status count cells
                            const val = ws[addr].v;
                            const hasData = val && val !== '-' && val !== 0;
                            ws[addr].s = {
                                fill: { patternType: 'solid', fgColor: { rgb: rowBg } },
                                font: { color: { rgb: hasData ? 'E2E8F0' : '475569' }, bold: hasData, sz: 10, name: 'Calibri' },
                                alignment: { horizontal: 'center', vertical: 'center' },
                                border: { top: { style: 'thin', color: { rgb: '334155' } }, bottom: { style: 'thin', color: { rgb: '334155' } }, left: { style: 'thin', color: { rgb: '334155' } }, right: { style: 'thin', color: { rgb: '334155' } } }
                            };
                        }
                    }
                }
            }

            XLSX.utils.book_append_sheet(wb, ws, 'Caller Report');
            XLSX.writeFile(wb, fileName);
            return;
        }

        // ---- OTHER TABS: Simple export ----
        let aoa = [], title = '', fileName = '';
        const cols = ['Connected', 'Disconnected', 'Not Connected', 'Interested', 'Not Interested', 'Follow-up', 'Closed', 'Not Called'];
        let grandCount = 0, grandValue = 0;
        let colTotals = {};
        cols.forEach(c => colTotals[c] = 0);

        if (tab === 'main') {
            if (!window._wosgMainData) return;
            title = 'WITHOUT OSG MAIN REPORT â€”Â ' + dateStr;
            fileName = 'Without_OSG_Main_Report_' + dateStr + '.xlsx';
            aoa.push([title]);
            aoa.push(['SUMMARY', ...cols, 'TOTAL VALUE', 'TOTAL COUNT']);

            window._wosgMainData.forEach(grp => {
                let counts = {};
                grp.rows.forEach(r => counts[r.status] = r.count);

                let row = ['Overall'];
                cols.forEach(c => {
                    row.push(counts[c] || 0);
                    colTotals[c] += (counts[c] || 0);
                });
                row.push(Math.round(grp.totalValue));
                row.push(grp.totalCount);
                aoa.push(row);

                grandCount += grp.totalCount; grandValue += grp.totalValue;
            });
        } else if (tab === 'caller') {
            if (!window._wosgCallerData || !window._wosgCallers) return;
            title = 'WITHOUT OSG CALLER REPORT â€”Â ' + dateStr;
            fileName = 'Without_OSG_Caller_Report_' + dateStr + '.xlsx';
            aoa.push([title]);
            aoa.push(['CALLER', ...cols, 'TOTAL VALUE', 'TOTAL COUNT']);

            window._wosgCallers.forEach(callerName => {
                const cd = window._wosgCallerData[callerName];
                let counts = {};
                cd.rows.forEach(r => counts[r.status] = r.count);

                let row = [callerName];
                cols.forEach(c => {
                    row.push(counts[c] || 0);
                    colTotals[c] += (counts[c] || 0);
                });
                row.push(Math.round(cd.totalValue));
                row.push(cd.totalCount);
                aoa.push(row);

                grandCount += cd.totalCount; grandValue += cd.totalValue;
            });
        } else if (tab === 'daily') {
            if (!window._wosgDailyData) return;
            title = 'WITHOUT OSG DAILY REPORT â€”Â ' + dateStr;
            fileName = 'Without_OSG_Daily_Report_' + dateStr + '.xlsx';
            aoa.push([title]);
            aoa.push(['DATE', ...cols, 'TOTAL VALUE', 'TOTAL COUNT']);

            window._wosgDailyData.forEach(grp => {
                let counts = {};
                grp.rows.forEach(r => counts[r.status] = r.count);

                let row = [grp.date];
                cols.forEach(c => {
                    row.push(counts[c] || 0);
                    colTotals[c] += (counts[c] || 0);
                });
                row.push(Math.round(grp.totalValue));
                row.push(grp.totalCount);
                aoa.push(row);

                grandCount += grp.totalCount; grandValue += grp.totalValue;
            });
        } else if (tab === 'monthly') {
            if (!window._wosgMonthlyData) return;
            title = 'WITHOUT OSG MONTHLY REPORT â€”Â ' + dateStr;
            fileName = 'Without_OSG_Monthly_Report_' + dateStr + '.xlsx';
            aoa.push([title]);
            aoa.push(['MONTH', ...cols, 'TOTAL VALUE', 'TOTAL COUNT']);

            window._wosgMonthlyData.forEach(grp => {
                let counts = {};
                grp.rows.forEach(r => counts[r.status] = r.count);

                let row = [grp.date];
                cols.forEach(c => {
                    row.push(counts[c] || 0);
                    colTotals[c] += (counts[c] || 0);
                });
                row.push(Math.round(grp.totalValue));
                row.push(grp.totalCount);
                aoa.push(row);

                grandCount += grp.totalCount; grandValue += grp.totalValue;
            });
        }

        let gtRow = ['GRAND TOTAL'];
        cols.forEach(c => gtRow.push(colTotals[c]));
        gtRow.push(Math.round(grandValue));
        gtRow.push(grandCount);
        aoa.push(gtRow);

        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(aoa);

        // Col widths based on pivot structure
        let wscols = [{ wch: 18 }]; // Caller/Date/Month
        cols.forEach(() => wscols.push({ wch: 14 }));
        wscols.push({ wch: 16 }); // Total Val
        wscols.push({ wch: 12 }); // Total Cnt
        ws['!cols'] = wscols;

        ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: cols.length + 2 } }];

        XLSX.utils.book_append_sheet(wb, ws, tab === 'caller' ? 'Caller Report' : (tab === 'daily' ? 'Daily Report' : (tab === 'main' ? 'Main Report' : 'Monthly Report')));
        XLSX.writeFile(wb, fileName);
    };


    // ---- CUSTOMERS WITHOUT OSG PAGE ----
    $('btnCORefresh').addEventListener('click', renderCustomersOSGPage);
    $('btnCOExport').addEventListener('click', exportCustomersOSGExcel);
    const btnDue = document.getElementById('btnCODueToday');
    if (btnDue) {
        btnDue.addEventListener('click', () => {
            window.coDueTodayFilter = !window.coDueTodayFilter;
            if (window.coDueTodayFilter) {
                btnDue.style.boxShadow = '0 0 0 3px rgba(239,68,68,0.5)';
                btnDue.innerHTML = 'ÃƒÂ¢Ã‚ÂÃ…â€™ Clear Due Today';
            } else {
                btnDue.style.boxShadow = '0 2px 8px rgba(239,68,68,0.3)';
                btnDue.innerHTML = 'ÃƒÂ°Ã…Â¸Ã¢â‚¬ÂÃ‚Â¥ Due Today';
            }
            renderCustomersOSGPage();
        });
    }
    document.querySelector('[data-section="customers-osg-section"]').addEventListener('click', async () => {
        // Build month switcher immediately from Firebase, then load statuses & render
        if (typeof updateMonthSwitcherUI === 'function') await updateMonthSwitcherUI();
        // Render page immediately so data shows without waiting for Firebase
        renderCustomersOSGPage();
        // Then load statuses from Firebase and re-render with real-time status data
        loadCoStatuses(() => setTimeout(renderCustomersOSGPage, 50));
    });
    ['coBrand', 'coRBM', 'coBDM', 'coProduct', 'coBranch', 'coSort', 'coStatusFilter', 'coCallerFilter'].forEach(id => {
        $(id).addEventListener('change', () => {
            if (id === 'coSort') window.coSortMode = $('coSort').value;
            renderCustomersOSGPage();
        });
    });

    let isCoStatusLive = false;
    let coLiveRef = null; // Track the active Firebase listener ref


    function loadCoStatuses(callback) {
        if (typeof firebase === 'undefined') { if (callback) callback(); return; }
        const month = window.coActiveMonth || new Date().toISOString().substring(0, 7);
        const path = 'customerStatus/' + month;

        if (!isCoStatusLive) {
            isCoStatusLive = true;

            // Detach any old listeners
            if (coLiveRef) {
                coLiveRef.off('value');
                coLiveRef.off('child_changed');
                coLiveRef.off('child_added');
            }
            coLiveRef = firebase.database().ref(path);
            coStatusMap = {};

            // ===== REAL-TIME LISTENER (attached FIRST, before initial load) =====
            // This ensures NO change event is ever missed
            coLiveRef.on('child_changed', childSnap => {
                const inv = childSnap.key;
                coStatusMap[fbKey(inv)] = childSnap.val();
                const st = coStatusMap[fbKey(inv)];

                // If another caller made ANY selection ÃƒÂ¢Ã¢â‚¬Â Ã¢â‚¬â„¢ instantly remove from this caller's view
                if (st.calledBy && st.calledBy !== currentCaller) {
                    const rowEl = document.getElementById('co-row-' + fbKey(inv));
                    if (rowEl) rowEl.remove();
                } else {
                    // Own row or unclaimed ÃƒÂ¢Ã¢â‚¬Â Ã¢â‚¬â„¢ update status colour in-place
                    const rowEl = document.getElementById('co-row-' + fbKey(inv));
                    if (rowEl && !rowEl.contains(document.activeElement)) {
                        const rowData = (coCurrentRows || []).find(r => fbKey(r.invoice) === inv);
                        if (rowData && typeof updateCoSingleRow === 'function') updateCoSingleRow(rowData.invoice);
                    }
                }

                if (typeof updateCoStatsInPlace === 'function') updateCoStatsInPlace();
            });

            // ===== INITIAL FULL LOAD (once) â€”Â then render page =====
            coLiveRef.once('value', snap => {
                const data = snap.val() || {};
                Object.keys(data).forEach(inv => { coStatusMap[fbKey(inv)] = data[inv]; });
                if (callback) { callback(); callback = null; }
            });

        } else {
            if (callback) callback();
        }
    }

    // Flashes a row with a colour pulse, shows who claimed it, then slides it away
    function _animateRowClaimed(rowEl, callerName) {
        const callerObj = (typeof CO_CALLERS !== 'undefined' ? CO_CALLERS : []).find(x => x.name === callerName);
        const claimColor = callerObj ? callerObj.color : '#f59e0b';
        const claimBg    = callerObj ? callerObj.bg    : 'rgba(245,158,11,0.18)';

        // Inject claim-flash CSS once
        if (!document.getElementById('_co_claim_style')) {
            const s = document.createElement('style');
            s.id = '_co_claim_style';
            s.textContent = `
                @keyframes co-claim-pulse {
                    0%   { opacity:1; }
                    30%  { opacity:0.9; }
                    60%  { opacity:1; }
                    100% { opacity:0; }
                }
                .co-row-claiming {
                    animation: co-claim-pulse 1.4s ease forwards !important;
                    pointer-events: none !important;
                    overflow: hidden;
                }
            `;
            document.head.appendChild(s);
        }

        // Overlay a banner showing who claimed it
        rowEl.style.position = 'relative';
        rowEl.style.background = claimBg;
        rowEl.style.outline = '2px solid ' + claimColor;
        rowEl.style.transition = 'all 0.3s';

        const banner = document.createElement('td');
        banner.colSpan = 10;
        banner.style.cssText = `position:absolute;inset:0;display:flex;align-items:center;justify-content:center;
            background:${claimBg};color:${claimColor};font-weight:700;font-size:0.85rem;gap:8px;border-radius:6px;`;
        banner.innerHTML = `<span style="font-size:1.1rem;">ÃƒÂ°Ã…Â¸Ã¢â‚¬ÂÃ¢â‚¬â„¢</span> Claimed by <strong>${callerName}</strong> â€”Â removing from your list...`;
        rowEl.innerHTML = '';
        rowEl.appendChild(banner);
        rowEl.classList.add('co-row-claiming');

        // Remove from DOM + coCurrentRows after animation
        setTimeout(() => {
            rowEl.style.maxHeight = (rowEl.offsetHeight) + 'px';
            rowEl.style.transition = 'max-height 0.35s ease, opacity 0.35s ease, padding 0.35s ease';
            requestAnimationFrame(() => {
                rowEl.style.maxHeight = '0';
                rowEl.style.opacity  = '0';
                rowEl.style.paddingTop = '0';
                rowEl.style.paddingBottom = '0';
            });
            setTimeout(() => {
                rowEl.remove();
                if (typeof updateCoStatsInPlace === 'function') updateCoStatsInPlace();
            }, 380);
        }, 1450);
    }

    // Switch to a different month's archive
    window.switchCoMonth = async function (month) {
        window.coActiveMonth = month;

        // Try to load local data
        const localData = await loadMonthlyDataFromDB(month);
        if (localData) {
            productData = localData.productData || [];
            osgData = localData.osgData || [];
            amcData = localData.amcData || [];
            samsungData = localData.samsungData || [];
            allData = [...productData, ...amcData];
            console.log('[IndexedDB] Loaded data for month:', month);
        } else {
            console.warn('[IndexedDB] No local data found for month', month);
            productData = [];
            osgData = [];
            amcData = [];
            samsungData = [];
            allData = [];
        }

        isCoStatusLive = false; // Force re-subscribe
        loadCoStatuses(() => renderCustomersOSGPage());
        if (typeof updateMonthSwitcherUI === 'function') updateMonthSwitcherUI();
    };

    // Build month switcher dropdown from Firebase keys + current month
    window.buildMonthSwitcher = async function () {
        if (typeof firebase === 'undefined') return;
        const localMonths = await getAllMonthsFromDB();
        return firebase.database().ref('customerStatus').once('value').then(snap => {
            let keys = Object.keys(snap.val() || {}).filter(k => /^\d{4}-\d{2}$/.test(k));
            localMonths.forEach(m => { 
                let norm = m;
                if (/^\d{4}-\d{1}$/.test(m)) norm = m.substring(0, 5) + '0' + m.substring(5);
                if (!keys.includes(norm) && /^\d{4}-\d{2}$/.test(norm)) keys.push(norm); 
            });
            keys.sort().reverse();

            const current = window.coActiveMonth || '';
            // Auto-set to the most recent month available in Firebase if not already set
            if (!window.coActiveMonth && keys.length > 0) {
                window.coActiveMonth = keys[0];
            }
            if (!window.coActiveMonth) {
                window.coActiveMonth = new Date().toISOString().substring(0, 7);
            }
            if (!keys.includes(window.coActiveMonth)) keys.unshift(window.coActiveMonth);
            const el = document.getElementById('coMonthSwitcher');
            if (!el) return;
            el.innerHTML = keys.map(k => {
                const [y, m] = k.split('-');
                const label = new Date(y, m - 1, 1).toLocaleString('default', { month: 'long', year: 'numeric' });
                const isLocal = localMonths.includes(k) ? ' (Local)' : '';
                return `<option value="${k}" ${k === window.coActiveMonth ? 'selected' : ''}>${label}${isLocal}</option>`;
            }).join('');
            // Update badge after dropdown is built
            const [y, m] = window.coActiveMonth.split('-');
            const label = new Date(y, m - 1, 1).toLocaleString('default', { month: 'long', year: 'numeric' });
            const badge = document.getElementById('coMonthBadge');
            if (badge) badge.textContent = label;
        });
    };

    // Update the month switcher UI badge
    window.updateMonthSwitcherUI = async function () {
        const month = window.coActiveMonth || new Date().toISOString().substring(0, 7);
        const [y, m] = month.split('-');
        const label = new Date(y, m - 1, 1).toLocaleString('default', { month: 'long', year: 'numeric' });
        const badge = document.getElementById('coMonthBadge');
        if (badge) badge.textContent = label;
        const el = document.getElementById('coMonthSwitcher');
        if (el) el.value = month;
        await window.buildMonthSwitcher();
    };

    // Sanitize invoice keys for Firebase (no . # $ [ ] / allowed)
    function fbKey(inv) {
        return String(inv).replace(/[.#$\[\]\/]/g, '_');
    }

    function saveCoStatus(inv) {
        if (typeof firebase === 'undefined') { console.error('[saveCoStatus] Firebase not loaded!'); return; }
        const month = window.coActiveMonth || new Date().toISOString().substring(0, 7);
        const status = coStatusMap[fbKey(inv)] || { callStatus: null, interest: null };
        const safeKey = fbKey(inv);
        
        console.log('[saveCoStatus] Saving:', month, safeKey, JSON.stringify(status));
        firebase.database().ref('customerStatus/' + month + '/' + safeKey).set(status)
            .then(() => { console.log('[saveCoStatus] SUCCESS:', safeKey); })
            .catch(e => {
                console.error('[saveCoStatus] FAILED:', safeKey, e);
                alert('Failed to save status to server: ' + e.message);
            });
    }





    function renderCustomersOSGPage() {
        let missedUnique = [];
        let sourceData = [];

        if (window.sharedMissedUnique) {
            // Restore original filters for shared view
            document.querySelectorAll('#customers-osg-section .lowconv-controls').forEach(el => el.style.display = 'flex');
            // Remove the custom shared filter bar if it exists
            const bar = document.getElementById('coSharedFilterBar');
            if (bar) bar.remove();

            sourceData = window.sharedMissedUnique;
        } else {
            // Standard dynamic processing
            if (productData.length === 0) {
                $('coMissedTable').innerHTML = noDataHTML('Upload data and generate reports first.');
                $('coMissedCount').textContent = '0 customers';
                $('coMissedCount').style.background = '';
                return;
            }
            sourceData = productData;
        }

        const selBrand = document.getElementById('coBrand') ? document.getElementById('coBrand').value : '';
        const selRBM = $('coRBM').value;
        const selBDM = $('coBDM').value;
        const selProduct = $('coProduct').value;
        const selBranch = $('coBranch').value;
        const selDate = document.getElementById('coDateFilter') ? document.getElementById('coDateFilter').value : '';
        coCurrentSelDate = selDate;
        const selStatusFilter = document.getElementById('coStatusFilter') ? document.getElementById('coStatusFilter').value : '';
        const selCallerFilter = document.getElementById('coCallerFilter') ? document.getElementById('coCallerFilter').value : '';

        // Populate filter dropdowns (preserve selection)
        const brandSet = [...new Set(sourceData.map(r => r.brand).filter(Boolean))].sort();
        const rbmSet = [...new Set(sourceData.map(r => r.rbm).filter(Boolean))].sort();
        const bdmSet = [...new Set(sourceData.map(r => r.bdm).filter(Boolean))].sort();
        const prodSet = [...new Set(sourceData.map(r => r.product).filter(Boolean))].sort();
        const branchSet = [...new Set(sourceData.map(r => r.branch).filter(Boolean))].sort();

        const coBrandEl = document.getElementById('coBrand');
        if (coBrandEl) { coBrandEl.innerHTML = '<option value="">All Brands</option>' + brandSet.map(b => `<option value="${b}">${b}</option>`).join(''); coBrandEl.value = selBrand; }
        const coRBMEl = document.getElementById('coRBM');
        if (coRBMEl) { coRBMEl.innerHTML = '<option value="">All RBMs</option>' + rbmSet.map(r => `<option value="${r}">${r}</option>`).join(''); coRBMEl.value = selRBM; }
        const coBDMEl = document.getElementById('coBDM');
        if (coBDMEl) { coBDMEl.innerHTML = '<option value="">All BDMs</option>' + bdmSet.map(b => `<option value="${b}">${b}</option>`).join(''); coBDMEl.value = selBDM; }
        const coProductEl = document.getElementById('coProduct');
        if (coProductEl) { coProductEl.innerHTML = '<option value="">All Products</option>' + prodSet.map(p => `<option value="${p}">${p}</option>`).join(''); coProductEl.value = selProduct; }
        const coBranchEl = document.getElementById('coBranch');
        if (coBranchEl) { coBranchEl.innerHTML = '<option value="">All Branches</option>' + branchSet.map(b => `<option value="${b}">${b}</option>`).join(''); coBranchEl.value = selBranch; }

        // Filter product rows
        let filtP = sourceData;
        if (selBrand) filtP = filtP.filter(r => r.brand === selBrand);
        if (selRBM) filtP = filtP.filter(r => r.rbm === selRBM);
        if (selBDM) filtP = filtP.filter(r => r.bdm === selBDM);
        if (selProduct) filtP = filtP.filter(r => r.product === selProduct);
        if (selBranch) filtP = filtP.filter(r => r.branch === selBranch);

        if (selDate) {
            filtP = filtP.filter(r => {
                let dStr = '';
                let dt = r.invoiceDate || r.time;
                if (!dt) return false;

                if (dt instanceof Date && !isNaN(dt)) {
                    const y = dt.getFullYear();
                    const m = String(dt.getMonth() + 1).padStart(2, '0');
                    const d = String(dt.getDate()).padStart(2, '0');
                    dStr = `${y}-${m}-${d}`;
                } else if (typeof dt === 'string') {
                    if (dt.match(/^\d{4}-\d{2}-\d{2}/)) {
                        dStr = dt.substring(0, 10);
                    } else {
                        let match = dt.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
                        if (match) {
                            let yy = match[3].length === 2 ? "20" + match[3] : match[3];
                            dStr = `${yy}-${match[2].padStart(2, '0')}-${match[1].padStart(2, '0')}`;
                        } else {
                            const pd = new Date(dt);
                            if (!isNaN(pd)) {
                                dStr = `${pd.getFullYear()}-${String(pd.getMonth() + 1).padStart(2, '0')}-${String(pd.getDate()).padStart(2, '0')}`;
                            }
                        }
                    }
                } else if (typeof dt === 'number') {
                    // Excel serial number (days since 1900)
                    let dObj = new Date(Math.round((dt - 25569) * 86400 * 1000));
                    dStr = `${dObj.getUTCFullYear()}-${String(dObj.getUTCMonth() + 1).padStart(2, '0')}-${String(dObj.getUTCDate()).padStart(2, '0')}`;
                }
                return dStr === selDate;
            });
        }

        // Build OSG invoice set (empty if none uploaded, harmless in shared view)
        const osgInvoices = new Set();
        (typeof osgData !== 'undefined' ? osgData : []).forEach(r => { if (r.invoice) osgInvoices.add(r.invoice); });

        // Find product rows with no matching OSG invoice, deduplicated by invoice
        const seenInv = new Set();
        filtP.forEach(r => {
            // In shared view, sourceData is already osg-filtered
            const isMissed = window.sharedMissedUnique ? true : (r.invoice && !osgInvoices.has(r.invoice));
            if (isMissed && !seenInv.has(r.invoice)) {
                seenInv.add(r.invoice);
                missedUnique.push(r);
            }
        });

        // Hide rows already claimed by ANOTHER caller (they won't appear in this caller's list)
        // Only apply when a caller is logged in
        if (currentCaller) {
            const visible = missedUnique.filter(r => {
                const st = coStatusMap[fbKey(r.invoice)] || {};
                // Show row if: unclaimed, OR claimed by the current caller
                return !st.calledBy || st.calledBy === currentCaller;
            });
            missedUnique.length = 0;
            visible.forEach(r => missedUnique.push(r));
        }

        // Apply CRM Status filter dropdown
        if (selStatusFilter) {
            const statusFiltered = missedUnique.filter(r => {
                const st = coStatusMap[fbKey(r.invoice)] || {};
                if (selStatusFilter === 'connected')      return st.callStatus === 'connected';
                if (selStatusFilter === 'not-connected')  return st.callStatus === 'not-connected' || st.callStatus === 'disconnected';
                if (selStatusFilter === 'follow-up')      return st.interest === 'follow-up';
                if (selStatusFilter === 'interested')     return st.interest === 'interested';
                if (selStatusFilter === 'not-interested') return st.interest === 'not-interested';
                return true;
            });
            missedUnique.length = 0;
            statusFiltered.forEach(r => missedUnique.push(r));
        }

        // Apply manual caller filter dropdown (admin view â€”Â shows rows by a specific caller)
        if (selCallerFilter) {
            const filtered = missedUnique.filter(r => {
                const st = coStatusMap[fbKey(r.invoice)] || {};
                return st.calledBy === selCallerFilter;
            });
            missedUnique.length = 0;
            filtered.forEach(r => missedUnique.push(r));
        }

        if (window.coDueTodayFilter) {
            const today = new Date().toISOString().substring(0, 10);
            const filtered = missedUnique.filter(r => {
                const st = coStatusMap[fbKey(r.invoice)] || {};
                return st.interest === 'follow-up' && st.followUpDate && st.followUpDate <= today;
            });
            missedUnique.length = 0;
            filtered.forEach(r => missedUnique.push(r));
        }

        if (window.sharedMissedUnique) {
            $('coMissedCount').textContent = `${missedUnique.length} of ${window.sharedMissedUnique.length} customers (Shared View)`;
            $('coMissedCount').style.background = 'linear-gradient(135deg, #10b981, #059669)';
        } else {
            $('coMissedCount').textContent = `${missedUnique.length} customers`;
            $('coMissedCount').style.background = '';
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
            $('coMissedTable').innerHTML = noDataHTML('All invoices have OSG entries ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â great conversion! Ãƒâ€šÃ‚Â°');
            return;
        }

        // Store filtered rows globally for pagination and single-row updates
        coCurrentRows = missedUnique;
        coDisplayLimit = 100;

        // ---- Stats Bar (with IDs for in-place update) ----
        const total = missedUnique.length;
        const connected = missedUnique.filter(r => (coStatusMap[fbKey(r.invoice)] || {}).callStatus === 'connected').length;
        const disconnected = missedUnique.filter(r => (coStatusMap[fbKey(r.invoice)] || {}).callStatus === 'disconnected').length;
        const interested = missedUnique.filter(r => (coStatusMap[fbKey(r.invoice)] || {}).interest === 'interested').length;
        const notInterested = missedUnique.filter(r => (coStatusMap[fbKey(r.invoice)] || {}).interest === 'not-interested').length;
        const notCalled = total - connected - disconnected;

        const statsBar = `
        <div id="coStatsBar" style="display:flex; gap:12px; flex-wrap:wrap; margin-bottom:20px;">
            <div style="flex:1;min-width:120px;background:var(--bg-card);border:1px solid var(--border);border-radius:12px;padding:14px 18px;border-top:3px solid #64748b;">
                <div id="coStat_total" style="font-size:1.6rem;font-weight:700;color:var(--text-primary);">${total}</div>
                <div style="font-size:0.78rem;color:var(--text-muted);margin-top:2px;">Total Customers</div>
            </div>
            <div style="flex:1;min-width:120px;background:var(--bg-card);border:1px solid var(--border);border-radius:12px;padding:14px 18px;border-top:3px solid #64748b;">
                <div id="coStat_notCalled" style="font-size:1.6rem;font-weight:700;color:#94a3b8;">${notCalled}</div>
                <div style="font-size:0.78rem;color:var(--text-muted);margin-top:2px;">Not Yet Called</div>
            </div>
            <div style="flex:1;min-width:120px;background:var(--bg-card);border:1px solid var(--border);border-radius:12px;padding:14px 18px;border-top:3px solid #2563eb;">
                <div id="coStat_connected" style="font-size:1.6rem;font-weight:700;color:#2563eb;">${connected}</div>
                <div style="font-size:0.78rem;color:var(--text-muted);margin-top:2px;">Connected</div>
            </div>
            <div style="flex:1;min-width:120px;background:var(--bg-card);border:1px solid var(--border);border-radius:12px;padding:14px 18px;border-top:3px solid #9333ea;">
                <div id="coStat_disconnected" style="font-size:1.6rem;font-weight:700;color:#9333ea;">${disconnected}</div>
                <div style="font-size:0.78rem;color:var(--text-muted);margin-top:2px;">Disconnected</div>
            </div>
            <div style="flex:1;min-width:120px;background:var(--bg-card);border:1px solid var(--border);border-radius:12px;padding:14px 18px;border-top:3px solid #16a34a;">
                <div id="coStat_interested" style="font-size:1.6rem;font-weight:700;color:#16a34a;">${interested}</div>
                <div style="font-size:0.78rem;color:var(--text-muted);margin-top:2px;">Interested</div>
            </div>
            <div style="flex:1;min-width:120px;background:var(--bg-card);border:1px solid var(--border);border-radius:12px;padding:14px 18px;border-top:3px solid #dc2626;">
                <div id="coStat_notInterested" style="font-size:1.6rem;font-weight:700;color:#dc2626;">${notInterested}</div>
                <div style="font-size:0.78rem;color:var(--text-muted);margin-top:2px;">Not Interested</div>
            </div>
        </div>`;

        // ---- Caller Selector ----
        const callerSelector = `
        <div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin-bottom:20px;background:var(--bg-card);border:1px solid var(--border);border-radius:12px;padding:14px 18px;">
            <span style="display:inline-flex;align-items:center;gap:6px;font-size:0.85rem;font-weight:600;color:var(--text-muted);white-space:nowrap;"><svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></svg> You are:</span>
            <div style="display:flex;gap:8px;flex-wrap:wrap;">
                ${CO_CALLERS.map(c => `
                <button onclick="window.selectCoCaller('${c.name}')" style="
                    display:flex;align-items:center;gap:7px;
                    padding:7px 16px;border-radius:20px;border:2px solid ${c.color};
                    font-size:0.85rem;font-family:inherit;cursor:pointer;font-weight:600;
                    background:${currentCaller === c.name ? c.bg : 'transparent'};
                    color:${c.color};
                    box-shadow:${currentCaller === c.name ? '0 0 0 2px ' + c.color + '55' : 'none'};
                    transition:all 0.2s;
                ">
                    <span style="width:24px;height:24px;border-radius:50%;background:${c.color};color:#fff;display:inline-flex;align-items:center;justify-content:center;font-size:0.75rem;font-weight:700;">${c.name[0]}</span>
                    ${c.name}
                    ${currentCaller === c.name ? '<span style="font-size:0.7rem;opacity:0.8;">â€”Â Active</span>' : ''}
                </button>`).join('')}
            </div>
            ${currentCaller
                ? `<span style="margin-left:auto;font-size:0.8rem;color:var(--text-muted);">Logging calls as <strong style="color:var(--text-primary);">${currentCaller}</strong></span>`
                : '<span style="margin-left:auto;font-size:0.8rem;color:#f59e0b;"> Select your name to log calls</span>'}
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
            ? `<div id="coLoadMoreWrap" style="text-align:center;margin-top:16px;">
                <button onclick="window.coLoadMore()" style="
                    padding:10px 28px;border-radius:20px;border:1.5px solid var(--accent);
                    background:transparent;color:var(--accent);font-family:inherit;
                    font-size:0.88rem;font-weight:600;cursor:pointer;transition:all 0.2s;
                ">ÃƒÂ¢Ã‚Â¬Ã¢â‚¬Â¡ Load ${Math.min(remaining, 100)} more (${remaining} remaining)</button>
               </div>` : '';

        $('coMissedTable').innerHTML = statsBar + callerSelector + tableHeader + rowsHTML + '</tbody></table></div>' + loadMoreBtn;

        // ---- Handlers ----
        window.selectCoCaller = function (name) {
            // Find caller object to check password
            const callerObj = CO_CALLERS.find(c => c.name === name);
            if (!callerObj) return;

            // If already selected, do nothing or toggling off
            if (currentCaller === name) {
                currentCaller = null;
                localStorage.removeItem('co_caller');
                renderCustomersOSGPage();
                return;
            }

            const attempt = window.prompt('Enter password for ' + name + ':');
            if (attempt === callerObj.pass || attempt === 'admin123') { // admin override just in case
                currentCaller = name;
                localStorage.setItem('co_caller', name);
                renderCustomersOSGPage();
            } else if (attempt !== null) {
                alert('Incorrect password for ' + name);
            }
        };

        window.coLoadMore = function () {
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
                        font-size:0.88rem;font-weight:600;cursor:pointer;">ÃƒÂ¢Ã‚Â¬Ã¢â‚¬Â¡ Load ${Math.min(remaining2, 100)} more (${remaining2} remaining)</button>`;
                } else {
                    wrap.remove();
                }
            }
        };

        window.toggleCoCall = function (inv, status) {
            if (!currentCaller && status !== "") {
                alert('Please select your name (Harmiya / Aswathi / Shikha / Anjana) at the top of the page before logging a call.');
                updateCoSingleRow(inv);
                return;
            }
            if (!coStatusMap[fbKey(inv)]) coStatusMap[fbKey(inv)] = { callStatus: null, interest: null, calledBy: null, remarks: '' };
            coStatusMap[fbKey(inv)].callStatus = status === "" ? null : status;
            if (status !== "") coStatusMap[fbKey(inv)].calledBy = currentCaller;

            if (!coStatusMap[fbKey(inv)].callStatus && !coStatusMap[fbKey(inv)].interest && !coStatusMap[fbKey(inv)].remarks) {
                coStatusMap[fbKey(inv)].calledBy = null;
                delete coStatusMap[fbKey(inv)].timestamp;
            } else {
                coStatusMap[fbKey(inv)].timestamp = new Date().toISOString();
            }
            saveCoStatus(inv);
            updateCoSingleRow(inv);
            updateCoStatsInPlace();
        };

        window.toggleCoInterest = function (inv, status) {
            if (!currentCaller && status !== "") {
                alert('Please select your name first.');
                updateCoSingleRow(inv);
                return;
            }
            if (!coStatusMap[fbKey(inv)]) coStatusMap[fbKey(inv)] = { callStatus: null, interest: null, calledBy: null, remarks: '', followUpDate: '' };
            coStatusMap[fbKey(inv)].interest = status === "" ? null : status;
            if (status !== "") coStatusMap[fbKey(inv)].calledBy = currentCaller;

            if (!coStatusMap[fbKey(inv)].callStatus && !coStatusMap[fbKey(inv)].interest && !coStatusMap[fbKey(inv)].remarks) {
                coStatusMap[fbKey(inv)].calledBy = null;
                delete coStatusMap[fbKey(inv)].timestamp;
            } else {
                coStatusMap[fbKey(inv)].timestamp = new Date().toISOString();
            }
            saveCoStatus(inv);
            updateCoSingleRow(inv);
            updateCoStatsInPlace();
        };

        window.setCoFollowUpDate = function (inv, dateStr) {
            if (!coStatusMap[fbKey(inv)]) coStatusMap[fbKey(inv)] = { callStatus: null, interest: null, calledBy: null, remarks: '', followUpDate: '' };
            coStatusMap[fbKey(inv)].followUpDate = dateStr;

            if (!coStatusMap[fbKey(inv)].callStatus && !coStatusMap[fbKey(inv)].interest && !coStatusMap[fbKey(inv)].remarks && (!dateStr || dateStr === '')) {
                coStatusMap[fbKey(inv)].calledBy = null;
                delete coStatusMap[fbKey(inv)].timestamp;
            } else {
                coStatusMap[fbKey(inv)].timestamp = new Date().toISOString();
            }

            saveCoStatus(inv);
            updateCoSingleRow(inv);
        };

        window.saveCoRemark = function (inv, remark) {
            if (!coStatusMap[fbKey(inv)]) coStatusMap[fbKey(inv)] = { callStatus: null, interest: null, calledBy: null, remarks: '' };
            coStatusMap[fbKey(inv)].remarks = remark;

            if (!coStatusMap[fbKey(inv)].callStatus && !coStatusMap[fbKey(inv)].interest && !coStatusMap[fbKey(inv)].remarks) {
                coStatusMap[fbKey(inv)].calledBy = null;
                delete coStatusMap[fbKey(inv)].timestamp;
            } else {
                if (remark !== "" && currentCaller) {
                    // If they add a remark, claim the row if they are logged in
                    coStatusMap[fbKey(inv)].calledBy = currentCaller;
                }
                coStatusMap[fbKey(inv)].timestamp = new Date().toISOString();
            }
            saveCoStatus(inv);
            updateCoSingleRow(inv);
        };
    }

    // Updates just the 6 stat numbers in-place (no table re-render)
    function updateCoStatsInPlace() {
        const rows = coCurrentRows || [];
        const total = rows.length;
        const connected = rows.filter(r => (coStatusMap[fbKey(r.invoice)] || {}).callStatus === 'connected').length;
        const disconnected = rows.filter(r => (coStatusMap[fbKey(r.invoice)] || {}).callStatus === 'disconnected').length;
        const interested = rows.filter(r => (coStatusMap[fbKey(r.invoice)] || {}).interest === 'interested').length;
        const notInterested = rows.filter(r => (coStatusMap[fbKey(r.invoice)] || {}).interest === 'not-interested').length;
        const notCalled = total - connected - disconnected;

        const set = (id, val) => { const el = $(id); if (el) el.textContent = val; };
        set('coStat_total', total);
        set('coStat_notCalled', notCalled);
        set('coStat_connected', connected);
        set('coStat_disconnected', disconnected);
        set('coStat_interested', interested);
        set('coStat_notInterested', notInterested);

        // Also update Caller Performance Cards
        const cpEl = $('coCallerPerformance');
        if (cpEl) cpEl.innerHTML = renderCallerPerformanceHTML();
    }

    function renderCallerPerformanceHTML() {
        if (!CO_CALLERS || CO_CALLERS.length === 0) return '';

        let html = '<div class="caller-performance-grid">';
        CO_CALLERS.forEach(caller => {
            // Aggregate stats from global coStatusMap for THIS caller
            const stats = { conn: 0, disc: 0, int: 0, closed: 0, followup: 0 };
            Object.values(coStatusMap).forEach(st => {
                if (st.calledBy === caller.name) {
                    if (st.callStatus === 'connected') stats.conn++;
                    if (st.callStatus === 'disconnected') stats.disc++;
                    if (st.interest === 'interested') stats.int++;
                    if (st.interest === 'bought') stats.closed++;
                    if (st.interest === 'follow-up') stats.followup++;
                }
            });

            html += `
            <div class="caller-card" style="border-top: 3px solid ${caller.color};">
                <div class="caller-card-header">
                    <div class="caller-avatar" style="background:${caller.color};">${caller.name[0]}</div>
                    <div class="caller-name">${caller.name}</div>
                    <div style="margin-left:auto; font-size:0.65rem; color:var(--text-muted); font-weight:700; text-transform:uppercase; letter-spacing:0.05em;">Caller</div>
                </div>
                <div class="caller-stats-mini-grid">
                    <div class="mini-stat conn">
                        <span class="mini-stat-label">Connected</span>
                        <span class="mini-stat-val">${stats.conn}</span>
                    </div>
                    <div class="mini-stat int">
                        <span class="mini-stat-label">Interested</span>
                        <span class="mini-stat-val">${stats.int}</span>
                    </div>
                    <div class="mini-stat closed">
                        <span class="mini-stat-label">Closed</span>
                        <span class="mini-stat-val">${stats.closed}</span>
                    </div>
                    <div class="mini-stat followup">
                        <span class="mini-stat-label">Follow-up</span>
                        <span class="mini-stat-val">${stats.followup}</span>
                    </div>
                </div>
            </div>`;
        });
        html += '</div>';
        return html;
    }

    // Replaces a single row in-place without touching the rest of the table
    function updateCoSingleRow(inv) {
        const rowEl = document.getElementById('co-row-' + fbKey(inv));
        if (!rowEl) return;
        const idx = coCurrentRows.findIndex((r, i) => (r.invoice || String(i)) === inv);
        if (idx === -1) return;

        // If this row was just claimed by another caller in shared view, animate it out
        const st = coStatusMap[fbKey(inv)] || {};
        if (currentCaller && st.calledBy && st.calledBy !== currentCaller) {
            _animateRowClaimed(rowEl, st.calledBy);
            return;
        }

        const newRowHTML = buildCoRowHTML(coCurrentRows[idx], idx);
        rowEl.outerHTML = newRowHTML;
    }

    // Builds HTML for a single table row
    function buildCoRowHTML(r, i) {
        const inv = r.invoice || String(i);
        const st = coStatusMap[fbKey(inv)] || { callStatus: null, interest: null, calledBy: null, remarks: '' };

        const rowBg = st.interest === 'bought' ? 'rgba(16, 185, 129, 0.15)'
            : st.interest === 'interested' ? 'rgba(22, 163, 74, 0.08)'
                : st.interest === 'not-interested' ? 'rgba(220, 38, 38, 0.08)'
                    : st.callStatus === 'disconnected' ? 'rgba(147, 51, 234, 0.05)'
                        : st.callStatus === 'connected' ? 'rgba(59, 130, 246, 0.05)'
                            : 'transparent';

        const isFinalState = ['interested', 'not-interested', 'bought'].includes(st.interest) || st.callStatus === 'disconnected';
        const rowOpacity = isFinalState ? '0.55' : '1';

        // Locking logic
        const isLocked = st.calledBy && st.calledBy !== currentCaller;
        const disAttr = isLocked ? 'disabled' : '';
        const opcStyle = isLocked ? 'opacity:0.5; cursor:not-allowed;' : '';

        // Added 'follow-up' and 'bought' options to Interest
        const interestBtnOptions = `
            <option value="" ${!st.interest ? 'selected' : ''}>- Select -</option>
            <option value="interested" ${st.interest === 'interested' ? 'selected' : ''}>ÃƒÂ¢Ã…â€œÃ¢â‚¬Â¦ Interested</option>
            <option value="not-interested" ${st.interest === 'not-interested' ? 'selected' : ''}>ÃƒÂ¢Ã‚ÂÃ…â€™ Not Interested</option>
            <option value="follow-up" ${st.interest === 'follow-up' ? 'selected' : ''}>Follow-up</option>
            <option value="bought" ${st.interest === 'bought' ? 'selected' : ''}>Closed</option>
        `;

        let setDateBtn = '';
        if (st.interest === 'follow-up') {
            const fDate = st.followUpDate ? new Date(st.followUpDate).toLocaleDateString('en-GB') : 'Set date';
            setDateBtn = `<button ${disAttr} style="margin-top:6px; padding:6px 10px; border-radius:6px; border:1px solid rgba(255,255,255,0.1); background:#111827; color:#d1d5db; font-size:0.75rem; font-family:inherit; cursor:pointer; width:100%; display:flex; align-items:center; gap:6px; ${opcStyle}" onclick="window.coCalWidget.open(this, '${st.followUpDate || ''}', (val) => window.setCoFollowUpDate('${inv}', val))"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg> ${fDate}</button>`;
        }

        const interestBtn = `
            <div style="display:flex;flex-direction:column;max-width:140px;">
                <select ${disAttr} onchange="window.toggleCoInterest('${inv}', this.value)" style="
                    padding:6px 10px;border-radius:8px;border:1px solid ${st.interest === 'interested' ? '#16a34a' : (st.interest === 'not-interested' ? '#dc2626' : (st.interest === 'follow-up' ? '#f59e0b' : 'var(--border)'))};
                    font-size:0.85rem;font-family:inherit;cursor:pointer;font-weight:600;
                    background:var(--bg-input); color:var(--text-primary); outline:none; width:100%; ${opcStyle}">
                ${interestBtnOptions}
                </select>
                ${setDateBtn}
            </div>
        `;

        const callerInfo = st.calledBy ? (() => {
            const c = CO_CALLERS.find(x => x.name === st.calledBy) || { color: '#64748b', bg: 'rgba(100,116,139,0.15)' };
            return `<span style="display:inline-flex;align-items:center;gap:4px;padding:2px 8px;border-radius:10px;background:${c.bg};color:${c.color};font-size:0.7rem;font-weight:700;margin-top:5px;">
                <span style="width:14px;height:14px;border-radius:50%;background:${c.color};color:#fff;display:inline-flex;align-items:center;justify-content:center;font-size:0.55rem;font-weight:800;">${st.calledBy[0]}</span>
                ${st.calledBy}</span>`;
        })() : '';

        const callBtns = r.customerNo ? `
            <div style="display:flex;flex-direction:column;gap:5px;margin-top:6px;align-items:flex-start;">
                <select ${disAttr} onchange="window.toggleCoCall('${inv}', this.value)" style="
                    padding:6px 10px;border-radius:8px;border:1px solid ${st.callStatus === 'connected' ? '#2563eb' : (st.callStatus === 'disconnected' ? '#9333ea' : 'var(--border)')};
                    font-size:0.85rem;font-family:inherit;cursor:pointer;font-weight:600;
                    background:var(--bg-input); color:var(--text-primary); outline:none; max-width:140px; ${opcStyle}">
                    <option value="" ${!st.callStatus ? 'selected' : ''}>- Status -</option>
                    <option value="connected" ${st.callStatus === 'connected' ? 'selected' : ''}>Connected</option>
                    <option value="not-connected" ${st.callStatus === 'not-connected' ? 'selected' : ''}>ÃƒÂ°Ã…Â¸Ã¢â‚¬ÂÃ‚Â´ Not Connected</option>
                    <option value="disconnected" ${st.callStatus === 'disconnected' ? 'selected' : ''}>ÃƒÂ°Ã…Â¸Ã¢â‚¬Å“Ã‚Âµ Disconnected</option>
                </select>
                ${callerInfo}
            </div>` : '';

        const remarksInput = `
            <input ${disAttr} type="text" value="${q(st.remarks || '')}" onchange="window.saveCoRemark('${inv}', this.value)" placeholder="Remarks..." style="
                padding:6px 10px;border-radius:8px;border:1px solid var(--border);
                font-size:0.85rem;font-family:inherit;background:var(--bg-input);color:var(--text-primary);
                width:130px; outline:none; ${opcStyle}" />
        `;


        let dStr = '';
        if (st.timestamp) {
            // Always show caller activity timestamp
            const upd = new Date(st.timestamp);
            dStr = `<span style="color:#10b981;font-weight:600;" title="Status Updated">&#x2714; ${upd.toLocaleString('en-GB', { day: 'numeric', month: 'short', hour: 'numeric', minute: 'numeric' })}</span>`;
        } else if (coCurrentSelDate) {
            // Only show purchase date when a date filter is active
            let dt = r.time || r.invoiceDate;
            if (dt) {
                const pd = new Date(dt);
                if (!isNaN(pd)) {
                    dStr = pd.toLocaleString('en-GB', { day: 'numeric', month: 'short', year: 'numeric', hour: 'numeric', minute: 'numeric' });
                } else {
                    dStr = String(dt).substring(0, 16);
                }
            }
        }

        return `<tr id="co-row-${fbKey(inv)}" style="border-bottom:1px solid var(--border);background:${rowBg};opacity:${rowOpacity};transition:background 0.2s, opacity 0.2s;">
            <td style="padding:12px 10px;color:var(--text-muted);font-size:0.8rem;">${i + 1}</td>
            <td style="padding:12px 10px;font-family:monospace;font-size:0.82rem;color:var(--text-secondary); width:120px;">${dStr}</td>
            <td style="padding:12px 10px;"><strong style="color:var(--text-primary);font-size:0.9rem;">${r.customerName || '-'}</strong><div style="font-size:0.75rem;color:var(--text-muted)">${r.invoice}</div></td>
            <td style="padding:12px 10px; width:150px;">${interestBtn}</td>
            <td style="padding:12px 10px; width:160px;">
                <div style="display:flex;align-items:center;gap:6px;">
                    <span style="font-weight:600;color:var(--text-primary);font-size:0.88rem;">${r.customerNo || '-'}</span>
                    ${r.customerNo ? `<a href="tel:${r.customerNo}" title="Call" style="color:var(--primary);display:flex;padding:5px;border-radius:50%;background:rgba(59,130,246,0.12);"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72 12.84 12.84 0 0 0 .7 2.81 2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45 12.84 12.84 0 0 0 2.81.7A2 2 0 0 1 22 16.92z"/></svg></a>` : ''}
                    ${r.customerNo ? `<a href="${(function () {
                const phone = '91' + r.customerNo.replace(/\D/g, '');
                const name = (r.customerName || 'Customer').split(' ')[0];
                const prod = r.product || 'your product';
                const inv = r.invoice || '';
                const val = r.soldPrice && r.soldPrice > 0 ? ' (worth \u20B9' + r.soldPrice.toLocaleString('en-IN') + ')' : '';
                const msg = 'Dear ' + name + ',\n\nGreetings from myG ÃƒÂ°Ã…Â¸Ã‹Å“Ã…Â \n\nThank you for your recent purchase of *' + prod + '*' + val + ' (Invoice: ' + inv + ').\n\n We noticed your purchase does not yet include an *OSG Extended Warranty* plan. OSG covers:\n\nÃƒÂ¢Ã…â€œÃ¢â‚¬Â¦ Extended protection beyond manufacturer warranty\nÃƒÂ¢Ã…â€œÃ¢â‚¬Â¦ Free doorstep repair service\nÃƒÂ¢Ã…â€œÃ¢â‚¬Â¦ Zero hidden charges\nÃƒÂ¢Ã…â€œÃ¢â‚¬Â¦ Instant claim processing\n\nSecuring your device takes just a minute â€”Â and gives you complete peace of mind! \n\nWould you be interested? Reply *YES* and we will take care of everything.\n\nWarm regards,\nmyG Team';
                return 'https://wa.me/' + phone + '?text=' + encodeURIComponent(msg);
            })()}" target="_blank" title="WhatsApp (English)" style="color:#25D366;display:flex;padding:5px;border-radius:50%;background:rgba(37,211,102,0.12);"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 11.5a8.38 8.38 0 0 1-.9 3.8 8.5 8.5 0 0 1-7.6 4.7 8.38 8.38 0 0 1-3.8-.9L3 21l1.9-5.7a8.38 8.38 0 0 1-.9-3.8 8.5 8.5 0 0 1 4.7-7.6 8.38 8.38 0 0 1 3.8-.9h.5a8.48 8.48 0 0 1 8 8v.5z"/></svg></a>` : ''}
                    ${r.customerNo ? `<a href="${(function () {
                const phone = '91' + r.customerNo.replace(/\D/g, '');
                const name = (r.customerName || 'Customer').split(' ')[0];
                const prod = r.product || 'your product';
                const inv = r.invoice || '';
                const priceStr = r.soldPrice && r.soldPrice > 0 ? ' (\u20B9' + r.soldPrice.toLocaleString('en-IN') + ')' : '';
                const msgML = '\u0d2a\u0d4d\u0d30\u0d3f\u0d2f ' + name + ',\n\nmyG-\u0d7d \u0d28\u0d3f\u0d28\u0d4d\u0d28\u0d41\u0d33\u0d4d\u0d33 \u0d06\u0d36\u0d02\u0d38\u0d15\u0d7e\n\n\u0d28\u0d3f\u0d19\u0d4d\u0d19\u0d7e \u0d05\u0d1f\u0d41\u0d24\u0d4d\u0d24\u0d3f\u0d1f\u0d46 \u0d35\u0d3e\u0d19\u0d4d\u0d19\u0d3f\u0d2f ' + prod + priceStr + ' \u0d28\u0d4d \u0d28\u0d28\u0d4d\u0d26\u0d3f. (\u0d07\u0d7b\u0d35\u0d4b\u0d2f\u0d4d\u0d38\u0d4d: ' + inv + ').\n\n \u0d28\u0d3f\u0d19\u0d4d\u0d19\u0d33\u0d41\u0d1f\u0d46 \u0d2a\u0d7c\u0d1a\u0d4d\u0d1a\u0d47\u0d38\u0d3f\u0d7d \u0d07\u0d24\u0d41\u0d35\u0d30\u0d46 OSG \u0d0e\u0d15\u0d4d\u0d38\u0d4d\u0d31\u0d4d\u0d1f\u0d7b\u0d21\u0d21\u0d4d \u0d35\u0d3e\u0d31\u0d7b\u0d4d\u0d31\u0d3f \u0d2a\u0d4d\u0d32\u0d3e\u0d7b \u0d09\u0d7e\u0d2a\u0d4d\u0d2a\u0d46\u0d1f\u0d41\u0d24\u0d4d\u0d24\u0d3f\u0d2f\u0d3f\u0d1f\u0d4d\u0d1f\u0d3f\u0d32\u0d4d\u0d32 \u0d0e\u0d28\u0d4d\u0d28\u0d4d \u0d1e\u0d19\u0d4d\u0d19\u0d7e \u0d36\u0d4d\u0d30\u0d26\u0d4d\u0d27\u0d3f\u0d1a\u0d4d\u0d1a\u0d41. OSG \u0d35\u0d34\u0d3f \u0d28\u0d3f\u0d19\u0d4d\u0d19\u0d7e\u0d15\u0d4d\u0d15\u0d4d \u0d32\u0d2d\u0d3f\u0d15\u0d4d\u0d15\u0d41\u0d28\u0d4d\u0d28\u0d24\u0d4d:\n\u0d28\u0d3f\u0d19\u0d4d\u0d19\u0d7e\u0d15\u0d4d\u0d15\u0d4d \u0d24\u0d3e\u0d32\u0d4d\u0d2a\u0d30\u0d4d\u0d2f\u0d2e\u0d41\u0d23\u0d4d\u0d1f\u0d4b? YES \u0d0e\u0d28\u0d4d\u0d28\u0d4d \u0d2e\u0d31\u0d41\u0d2a\u0d1f\u0d3f \u0d28\u0d7d\u0d15\u0d41\u0d15, \u0d2c\u0d3e\u0d15\u0d4d\u0d15\u0d3f \u0d15\u0d3e\u0d30\u0d4d\u0d2f\u0d19\u0d4d\u0d19\u0d7e \u0d1e\u0d19\u0d4d\u0d19\u0d7e \u0d1a\u0d46\u0d2f\u0d4d\u0d24\u0d41 \u0d24\u0d30\u0d3e\u0d02.\n\n\u0d38\u0d4d\u0d28\u0d47\u0d39\u0d24\u0d4d\u0d24\u0d4b\u0d1f\u0d46,\n\u0d1f\u0d40\u0d02 myG';
                return 'https://wa.me/' + phone + '?text=' + encodeURIComponent(msgML);
            })()}" target="_blank" title="WhatsApp (Malayalam)" style="color:#25D366;display:flex;padding:3px 6px;border-radius:12px;background:rgba(37,211,102,0.12);font-size:0.75rem;font-weight:700;text-decoration:none;align-items:center;">ML</a>` : ''}
                </div>
                ${callBtns}
            </td>
            <td style="padding:12px 10px;">${remarksInput}</td>
            <td style="padding:12px 10px;color:var(--text-secondary);font-size:0.85rem;">${r.branch || '-'}</td>
            <td style="padding:12px 10px;color:var(--text-secondary);font-size:0.85rem;">${r.product || '-'}</td>
            <td style="padding:12px 10px;text-align:right;font-weight:600;color:var(--text-primary);font-size:0.88rem;white-space:nowrap;">${fmtShortHtml(Math.abs(r.soldPrice || 0))}</td>
        </tr>`;
    }
    function exportCustomersOSGExcel() {
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
        amcData.forEach(r => { if (r.invoice) osgInvoices.add(r.invoice); });
        samsungData.forEach(r => { if (r.invoice) osgInvoices.add(r.invoice); });

        const seenInv = new Set();
        const missedUnique = [];
        filtP.forEach(r => {
            if (r.invoice && !osgInvoices.has(r.invoice) && !seenInv.has(r.invoice)) {
                seenInv.add(r.invoice);
                missedUnique.push(r);
            }
        });

        if (missedUnique.length === 0) return;
        const hdr = ['#', 'Invoice No', 'Date', 'Customer Name', 'Customer No', 'Call Status', 'Interest', 'Follow-up Date', 'Called By', 'Remarks', 'Branch', 'Product', 'Sold Price'];
        const data = missedUnique.map((r, i) => {
            const st = coStatusMap[fbKey(r.invoice)] || {};
            let dStr = '';
            if (r.invoiceDate) dStr = new Date(r.invoiceDate).toLocaleDateString();
            else if (r.time) dStr = new Date(r.time).toLocaleDateString();

            return [
                i + 1, r.invoice, dStr, r.customerName || '', r.customerNo || '',
                st.callStatus || '', st.interest || '', st.followUpDate || '', st.calledBy || '', st.remarks || '',
                r.branch || '', r.product || '', Math.round(r.soldPrice || 0)
            ];
        });
        exportToStyledExcel(data, hdr, 'customers_without_osg.xlsx', 'Missed OSG Customers');
    }

    // ---- EXCEL UTILITIES (PROFESSIONAL) ----
    function exportToStyledExcel(dataBody, headers, filename, sheetName = 'Report') {
        const wb = XLSX.utils.book_new();
        const data = [headers, ...dataBody];
        const ws = XLSX.utils.aoa_to_sheet(data);

        const headerStyle = {
            font: { name: 'Calibri', sz: 11, bold: true, color: { rgb: 'FFFFFF' } },
            fill: { fgColor: { rgb: '1e293b' } },
            alignment: { horizontal: 'center', vertical: 'center' },
            border: {
                top: { style: 'thin', color: { rgb: '000000' } },
                bottom: { style: 'thin', color: { rgb: '000000' } },
                left: { style: 'thin', color: { rgb: '000000' } },
                right: { style: 'thin', color: { rgb: '000000' } }
            }
        };

        const cellStyle = {
            font: { name: 'Calibri', sz: 11 },
            border: {
                top: { style: 'thin', color: { rgb: 'CBD5E1' } },
                bottom: { style: 'thin', color: { rgb: 'CBD5E1' } },
                left: { style: 'thin', color: { rgb: 'CBD5E1' } },
                right: { style: 'thin', color: { rgb: 'CBD5E1' } }
            }
        };

        const altRowStyle = {
            ...cellStyle,
            fill: { fgColor: { rgb: 'F8FAFC' } }
        };

        for (let R = 0; R < data.length; ++R) {
            for (let C = 0; C < headers.length; ++C) {
                const cell_ref = XLSX.utils.encode_cell({ c: C, r: R });
                if (!ws[cell_ref]) ws[cell_ref] = { t: 's', v: '' };

                if (R === 0) {
                    ws[cell_ref].s = headerStyle;
                } else {
                    let baseStyle = R % 2 === 0 ? { ...altRowStyle } : { ...cellStyle };
                    const h = headers[C].toUpperCase();
                    
                    if (ws[cell_ref].t === 'n') {
                        let isPctCol = h.includes('%') || h.includes('CONV');
                        let isCurrencyCol = h.includes('PRICE') || h.includes('REVENUE') || h.includes('TAX') || h.includes('PROFIT') || h.includes('VALUE');
                        
                        baseStyle.alignment = { horizontal: 'center', vertical: 'center' };

                        if (isPctCol) {
                            baseStyle.z = '0"%"'; // No decimals
                            const val = ws[cell_ref].v;
                            let greenThresh = 10, yellowThresh = 5;

                            if (h.includes('VAL CONV')) {
                                greenThresh = 2; yellowThresh = 1;
                            } else if (h.includes('LG-AMC QTY CONV') || h.includes('SAMSUNG QTY CONV')) {
                                greenThresh = 15; yellowThresh = 10;
                            } else if (h.includes('OSG QTY CONV') || h.includes('QTY CONV')) {
                                greenThresh = 10; yellowThresh = 5;
                            }

                            if (val < yellowThresh) {
                                baseStyle = { ...baseStyle, font: { color: { rgb: "991B1B" }, bold: true }, fill: { fgColor: { rgb: "FEE2E2" } } };
                            } else if (val >= greenThresh) {
                                baseStyle = { ...baseStyle, font: { color: { rgb: "166534" }, bold: true }, fill: { fgColor: { rgb: "DCFCE7" } } };
                            } else {
                                baseStyle = { ...baseStyle, font: { color: { rgb: "9A3412" }, bold: true }, fill: { fgColor: { rgb: "FEF08A" } } };
                            }
                        } else if (isCurrencyCol) {
                            baseStyle.z = '"\u20B9"#,##0'; // No decimals
                        } else {
                            baseStyle.z = '#,##0'; // Commas, no decimals
                        }
                    } else {
                        if (h.includes('RANK') || h.includes('QTY') || h.includes('COUNT')) {
                            baseStyle.alignment = { horizontal: 'center', vertical: 'center' };
                        } else {
                            baseStyle.alignment = { vertical: 'center' };
                        }
                    }
                    ws[cell_ref].s = baseStyle;
                }
            }
        }

        ws['!cols'] = headers.map(h => ({ wch: Math.max(10, h.length + 5) }));
        XLSX.utils.book_append_sheet(wb, ws, sheetName.substring(0, 31));
        XLSX.writeFile(wb, filename);
    }

    // ---- DASHBOARD FILTERS & TABS ----
    function groupBy(arr, key) {
        const m = {};
        arr.forEach(r => { const k = r[key] || 'Unknown'; if (!m[k]) m[k] = []; m[k].push(r); });
        return m;
    }
    function formatCurrency(n) { return '\u20B9' + n.toLocaleString('en-IN', { maximumFractionDigits: 0 }); }
    function fmtShort(n) {
        if (Math.abs(n) >= 1e7) return '\u20B9' + (n / 1e7).toFixed(1) + 'Cr';
        if (Math.abs(n) >= 1e5) return '\u20B9' + (n / 1e5).toFixed(1) + 'L';
        if (Math.abs(n) >= 1e3) return '\u20B9' + (n / 1e3).toFixed(1) + 'K';
        return '\u20B9' + n.toFixed(0);
    }
            function fmtShortHtml(n) {
        return '<span title="\u20B9' + Number(n).toLocaleString('en-IN') + '" style="cursor:help; border-bottom:1px dotted #94a3b8;">' + fmtShort(n) + '</span>';
    }
    function formatNumber(n) { return n.toLocaleString('en-IN'); }
    function capitalize(s) { return s.charAt(0).toUpperCase() + s.slice(1); }
    function truncate(s, len) { return s.length > len ? s.substring(0, len) + '...' : s; }
    function showLoading(show) { loadingOverlay.classList.toggle('active', show); }
    function q(s) { return '"' + (s || '').replace(/"/g, '""') + '"'; }

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
        const osgInvoices = new Set([...osgData, ...amcData, ...samsungData].map(r => r.invoice).filter(Boolean));
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

        html += '<div style="margin-bottom: 16px;"><strong> High Volume, Low Conversion Branches:</strong><ul style="margin:8px 0 0 20px; color:var(--text-muted); line-height: 1.6;">';
        if (worstBranches.length) worstBranches.forEach(b => html += `<li><strong>${b.branch}</strong>: ${b.pQ} products sold but only ${b.oQ} OSG attached (${b.conv.toFixed(1)}%). Estimated missed tracking revenue: ${fmtShortHtml(b.valP * 0.1)}</li>`);
        else html += '<li>No significantly underperforming branches detected.</li>';
        html += '</ul></div>';

        html += '<div style="margin-bottom: 16px;"><strong>ÃƒÂ°Ã…Â¸Ã¢â‚¬â„¢Ã…Â½ Missed Premium Device Attachments:</strong><ul style="margin:8px 0 0 20px; color:var(--text-muted); line-height: 1.6;">';
        if (premiumMisses.length) premiumMisses.forEach(m => html += `<li><strong>${m.staff} (${m.branch})</strong> sold a ${m.product} for ${fmtShortHtml(m.soldPrice)} without OSG (Inv: ${m.invoice}).</li>`);
        else html += '<li>Great job! High-value premium products seem to be attached correctly.</li>';
        html += '</ul></div>';

        html += '<div><strong>ÃƒÂ°Ã…Â¸Ã¢â‚¬ËœÃ‚Â¤ Highest Opportunity Staff:</strong><ul style="margin:8px 0 0 20px; color:var(--text-muted); line-height: 1.6;">';
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
            reasonHTML += `<li><strong>Premium Pitch Avoidance:</strong> Found multiple premium devices > \u20B950K sold with no OSG attached. Sales reps might be avoiding the OSG pitch on high-ticket items out of fear of losing the primary sale due to total cart value shock.</li>`;
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
        if (premiumMisses.length) html += `<li><strong>Premium Bundling Rule:</strong> Institute a strict rule that any manager override/discount on products over \u20B950K ideally requires an OSG attachment commitment.</li>`;
        if (worstStaff.length) html += `<li><strong>Targeted PIPs:</strong> Place <strong>${worstStaff.map(s => `${s.staff}`).join(', ')}</strong> on an accelerated 7-day OSG pitch improvement plan.</li>`;
        html += '<li><strong>Daily Morning Brief:</strong> Have branch managers physically review the "Customers Without OSG" dashboard list from yesterday\'s data before the store opens to identify missed pitch opportunities and contact customers via the WhatsApp quick-links.</li>';
        html += '</ol>';
        html += '</div>';

        html += '</div>'; // End flex wrapper
        return html;
    }

    // ========================================================================
    // FUTURE STORES EXPORT DASHBOARD LOGIC
    // ========================================================================
    const modalFSDash = $('fsExportDashboardModal');
    const btnCloseFSDash = $('btnCloseFSDashboard');

    if (modalFSDash && btnCloseFSDash) {
        // Close modal
        btnCloseFSDash.addEventListener('click', () => {
            modalFSDash.style.display = 'none';
        });

        // Tab switching
        const dashTabs = document.querySelectorAll('.dash-tab');
        dashTabs.forEach(tab => {
            tab.addEventListener('click', (e) => {
                dashTabs.forEach(t => t.classList.remove('active'));
                document.querySelectorAll('.dash-tab-content').forEach(c => c.classList.remove('active'));
                e.currentTarget.classList.add('active');
                document.getElementById('dashTab-' + e.currentTarget.getAttribute('data-tab')).classList.add('active');
            });
        });

        // Filter change ÃƒÂ¢Ã¢â‚¬Â Ã¢â‚¬â„¢ re-render
        ['fsDashRBM', 'fsDashBDM', 'fsDashBranch', 'fsDashStaff'].forEach(id => {
            const el = $(id);
            if (el) el.addEventListener('change', () => renderFsExportDashboard(false));
        });
    }

    function renderFsExportDashboard(initFilters = false) {
        if (!productData || productData.length === 0) return;

        const selRBM = $('fsDashRBM').value;
        const selBDM = $('fsDashBDM').value;
        const selBranch = $('fsDashBranch').value;
        const selStaff = $('fsDashStaff').value;

        // Base filter for future stores
        let fProduct = productData.filter(r => r.branch && r.branch.toUpperCase().includes('FUTURE'));

        // Initialize Filters if needed
        if (initFilters) {
            const rSet = new Set(), bdmSet = new Set(), brSet = new Set(), stSet = new Set();
            fProduct.forEach(r => {
                if (r.rbm) rSet.add(r.rbm);
                if (r.bdm) bdmSet.add(r.bdm);
                if (r.branch) brSet.add(r.branch);
                if (r.staff) stSet.add(r.staff);
            });
            const popSel = (id, set, def) => {
                const el = $(id);
                if (!el || el.tagName !== 'SELECT') return;
                const curr = el.value;
                let h = `<option value="">${def}</option>`;
                Array.from(set).sort().forEach(v => h += `<option value="${v}">${v}</option>`);
                el.innerHTML = h;
                if (Array.from(set).includes(curr)) el.value = curr;
            };
            popSel('fsDashRBM', rSet, 'All RBMs');
            popSel('fsDashBDM', bdmSet, 'All BDMs');
            popSel('fsDashBranch', brSet, 'All Branches');
            popSel('fsDashStaff', stSet, 'All Staff');
        }

        // Apply Current Filters
        fProduct = fProduct.filter(r =>
            (!selRBM || r.rbm === selRBM) &&
            (!selBDM || r.bdm === selBDM) &&
            (!selBranch || r.branch === selBranch) &&
            (!selStaff || r.staff === selStaff)
        );

        // Map invoices only for Product grouping, not warranties
        const invMeta = {};
        fProduct.forEach(r => {
            if (r.invoice) invMeta[r.invoice] = { branch: r.branch, bdm: r.bdm || 'Unknown', staff: r.staff || 'Unknown' };
        });

        // Helper to resolve the correct branch entirely decoupled from product file invoices
        const resolveWarrMeta = (r) => {
            const m = window.getBranchMeta(r);
            if (!m && !r.branch) return null;
            return { 
                branch: m ? m.origBranch : r.branch, 
                bdm: r.bdm || (m ? m.bdm : ''), 
                rbm: r.rbm || (m ? m.rbm : ''), 
                staff: r.staff || '' 
            };
        };

        const filterWarrRow = r => {
            const meta = resolveWarrMeta(r);
            if (!meta || !meta.branch || !meta.branch.toUpperCase().includes('FUTURE')) return false;

            if (selBDM && meta.bdm !== selBDM) return false;
            if (selBranch && meta.branch !== selBranch) return false;
            if (selStaff && meta.staff !== selStaff) return false;
            // (Skipping RBM for dashboard UI to keep it simple, though it respects selRBM if added)
            return true;
        };

        const fOSG = osgData.filter(filterWarrRow);
        const fAMC = amcData.filter(filterWarrRow);
        const fSamsung = samsungData.filter(filterWarrRow);

        // ====================================================================
        // TAB 1: BRANCH OVERVIEW (Group by BDM -> Branch)
        // ====================================================================
        const brGrp = {}; // key: BDM|Branch
        fProduct.forEach(r => {
            const k = (r.bdm || 'Unknown') + '|' + r.branch;
            if (!brGrp[k]) brGrp[k] = { bdm: r.bdm || 'Unknown', branch: r.branch, p: [], o: [], a: [], s: [] };
            brGrp[k].p.push(r);
        });

        fOSG.forEach(r => {
            const m = resolveWarrMeta(r) || {};
            let bdm = m.bdm || 'Unknown';
            let branch = m.branch || 'Unknown';
            const k = bdm + '|' + branch;
            if (!brGrp[k]) brGrp[k] = { bdm, branch, p: [], o: [], a: [], s: [] };
            brGrp[k].o.push(r);
        });

        fAMC.forEach(r => {
            const m = resolveWarrMeta(r) || {};
            let bdm = m.bdm || 'Unknown';
            let branch = m.branch || 'Unknown';
            const k = bdm + '|' + branch;
            if (!brGrp[k]) brGrp[k] = { bdm, branch, p: [], o: [], a: [], s: [] };
            brGrp[k].a.push(r);
        });

        fSamsung.forEach(r => {
            const m = resolveWarrMeta(r) || {};
            let bdm = m.bdm || 'Unknown';
            let branch = m.branch || 'Unknown';
            const k = bdm + '|' + branch;
            if (!brGrp[k]) brGrp[k] = { bdm, branch, p: [], o: [], a: [], s: [] };
            brGrp[k].s.push(r);
        });

        let brHtml = `<tr><th>BDM</th><th>Branch</th><th>Prod Qty</th><th>OSG Qty</th><th>LG-AMC Qty</th><th>Samsung Qty</th><th>OSG Val Conv %</th><th>LG-AMC Val Conv %</th><th>Samsung Val Conv %</th></tr>`;
        Object.values(brGrp).sort((a, b) => a.bdm.localeCompare(b.bdm) || a.branch.localeCompare(b.branch)).forEach(grp => {
            const pQ = grp.p.reduce((s, r) => s + r.qty, 0);
            const oQ = grp.o.reduce((s, r) => s + r.qty, 0);
            const aQ = grp.a.reduce((s, r) => s + r.qty, 0);
            const sQ = grp.s.reduce((s, r) => s + r.qty, 0);
            const pR = grp.p.reduce((s, r) => s + r.soldPrice, 0);
            const oR = grp.o.reduce((s, r) => s + r.soldPrice, 0);
            const aR = grp.a.reduce((s, r) => s + r.soldPrice, 0);
            const sR = grp.s.reduce((s, r) => s + r.soldPrice, 0);

            const lgPR = grp.p.reduce((s, r) => s + ((r.brand && r.brand.toUpperCase().includes('LG')) ? r.soldPrice : 0), 0);
            const samsungAllowedCats = ['AC', 'MICROWAVE OVEN', 'REFRIGERATOR', 'WASHING MACHINE', 'TV'];
            const samPR = grp.p.reduce((s, r) => s + ((r.brand && r.brand.toUpperCase().includes('SAMSUNG') && r.product && samsungAllowedCats.includes(r.product.toUpperCase().trim())) ? r.soldPrice : 0), 0);

            const oConv = pR > 0 ? ((oR / pR) * 100).toFixed(2) : '0.00';
            const aConv = lgPR > 0 ? ((aR / lgPR) * 100).toFixed(2) : '0.00';
            const sConv = samPR > 0 ? ((sR / samPR) * 100).toFixed(2) : '0.00';

            brHtml += `<tr>
                <td>${grp.bdm}</td>
                <td><strong>${grp.branch}</strong></td>
                <td class="col-num">${pQ}</td>
                <td class="col-num">${oQ}</td>
                <td class="col-num">${aQ}</td>
                <td class="col-num">${sQ}</td>
                <td class="col-num">${oConv}%</td>
                <td class="col-num">${aConv}%</td>
                <td class="col-num">${sConv}%</td>
            </tr>`;
        });
        const tableBr = $('tableFsDashBranch');
        if (tableBr && tableBr.querySelector('tbody')) tableBr.querySelector('tbody').innerHTML = brHtml;

        // ====================================================================
        // TAB 2: STAFF OVERVIEW (Group by Branch -> Staff)
        // ====================================================================
        const stGrp = {}; // key: Branch|Staff
        fProduct.forEach(r => {
            const k = r.branch + '|' + (r.staff || 'Unknown');
            if (!stGrp[k]) stGrp[k] = { branch: r.branch, staff: r.staff || 'Unknown', p: [], o: [], a: [], s: [] };
            stGrp[k].p.push(r);
        });
        fOSG.forEach(r => {
            let branch = r.branch || ((r.invoice && invMeta[r.invoice]) ? invMeta[r.invoice].branch : 'Unknown');
            let staff = (r.invoice && invMeta[r.invoice]) ? invMeta[r.invoice].staff : 'Unknown';
            const k = branch + '|' + staff;
            if (!stGrp[k]) stGrp[k] = { branch, staff, p: [], o: [], a: [], s: [] };
            stGrp[k].o.push(r);
        });

        fAMC.forEach(r => {
            let branch = r.branch || ((r.invoice && invMeta[r.invoice]) ? invMeta[r.invoice].branch : 'Unknown');
            let staff = r.staff || ((r.invoice && invMeta[r.invoice]) ? invMeta[r.invoice].staff : 'Unknown');
            const k = branch + '|' + staff;
            if (!stGrp[k]) stGrp[k] = { branch, staff, p: [], o: [], a: [], s: [] };
            stGrp[k].a.push(r);
        });

        fSamsung.forEach(r => {
            let branch = r.branch || ((r.invoice && invMeta[r.invoice]) ? invMeta[r.invoice].branch : 'Unknown');
            let staff = (r.invoice && invMeta[r.invoice]) ? invMeta[r.invoice].staff : 'Unknown';
            const k = branch + '|' + staff;
            if (!stGrp[k]) stGrp[k] = { branch, staff, p: [], o: [], a: [], s: [] };
            stGrp[k].s.push(r);
        });

        let stHtml = `<tr><th>Branch</th><th>Staff</th><th>Prod Qty</th><th>OSG Qty</th><th>LG-AMC Qty</th><th>Samsung Qty</th><th>OSG Val Conv %</th><th>LG-AMC Val Conv %</th><th>Samsung Val Conv %</th></tr>`;
        Object.values(stGrp).sort((a, b) => a.branch.localeCompare(b.branch) || a.staff.localeCompare(b.staff)).forEach(grp => {
            const pQ = grp.p.reduce((s, r) => s + r.qty, 0);
            const oQ = grp.o.reduce((s, r) => s + r.qty, 0);
            const aQ = grp.a.reduce((s, r) => s + r.qty, 0);
            const sQ = grp.s.reduce((s, r) => s + r.qty, 0);
            const pR = grp.p.reduce((s, r) => s + r.soldPrice, 0);
            const oR = grp.o.reduce((s, r) => s + r.soldPrice, 0);
            const aR = grp.a.reduce((s, r) => s + r.soldPrice, 0);
            const sR = grp.s.reduce((s, r) => s + r.soldPrice, 0);

            const lgPR = grp.p.reduce((s, r) => s + ((r.brand && r.brand.toUpperCase().includes('LG')) ? r.soldPrice : 0), 0);
            const samsungAllowedCats = ['AC', 'MICROWAVE OVEN', 'REFRIGERATOR', 'WASHING MACHINE', 'TV'];
            const samPR = grp.p.reduce((s, r) => s + ((r.brand && r.brand.toUpperCase().includes('SAMSUNG') && r.product && samsungAllowedCats.includes(r.product.toUpperCase().trim())) ? r.soldPrice : 0), 0);

            const oConv = pR > 0 ? ((oR / pR) * 100).toFixed(2) : '0.00';
            const aConv = lgPR > 0 ? ((aR / lgPR) * 100).toFixed(2) : '0.00';
            const sConv = samPR > 0 ? ((sR / samPR) * 100).toFixed(2) : '0.00';

            stHtml += `<tr>
                <td>${grp.branch}</td>
                <td><strong>${grp.staff}</strong></td>
                <td class="col-num">${pQ}</td>
                <td class="col-num">${oQ}</td>
                <td class="col-num">${aQ}</td>
                <td class="col-num">${sQ}</td>
                <td class="col-num">${oConv}%</td>
                <td class="col-num">${aConv}%</td>
                <td class="col-num">${sConv}%</td>
            </tr>`;
        });
        const tableSt = $('tableFsDashStaff');
        if (tableSt && tableSt.querySelector('tbody')) tableSt.querySelector('tbody').innerHTML = stHtml;

        // ====================================================================
        // TAB 3: LG-AMC (Group by Category)
        // ====================================================================
        const amcProdGrp = {};
        const pLG = fProduct.filter(r => r.brand && r.brand.toUpperCase().includes('LG'));
        pLG.forEach(r => {
            const k = (r.product || 'Unknown').toUpperCase().trim();
            if (!amcProdGrp[k]) amcProdGrp[k] = { name: k, pQ: 0, pR: 0, aQ: 0, aR: 0 };
            amcProdGrp[k].pQ += r.qty || 0;
            amcProdGrp[k].pR += r.soldPrice || 0;
        });
        fAMC.forEach(r => {
            const k = (r.product || 'Unknown').toUpperCase().trim();
            if (!amcProdGrp[k]) amcProdGrp[k] = { name: k, pQ: 0, pR: 0, aQ: 0, aR: 0 };
            amcProdGrp[k].aQ += r.qty || 0;
            amcProdGrp[k].aR += r.soldPrice || 0;
        });

        let amcHtml = `<tr><th>LG Category</th><th>Prod Qty</th><th>AMC Qty</th><th>Prod Rev</th><th>AMC Rev</th><th>Qty Conv %</th><th>Val Conv %</th></tr>`;
        Object.values(amcProdGrp).sort((a, b) => b.pQ - a.pQ).forEach(grp => {
            const qConv = grp.pQ > 0 ? ((grp.aQ / grp.pQ) * 100).toFixed(2) : '0.00';
            const vConv = grp.pR > 0 ? ((grp.aR / grp.pR) * 100).toFixed(2) : '0.00';
            amcHtml += `<tr>
                <td><strong>${grp.name}</strong></td>
                <td class="col-num">${grp.pQ}</td>
                <td class="col-num">${grp.aQ}</td>
                <td class="col-num">${fmtShortHtml(grp.pR)}</td>
                <td class="col-num">${fmtShortHtml(grp.aR)}</td>
                <td class="col-num">${qConv}%</td>
                <td class="col-num">${vConv}%</td>
            </tr>`;
        });
        const tableLg = $('tableFsDashLgAmc');
        if (tableLg && tableLg.querySelector('tbody')) tableLg.querySelector('tbody').innerHTML = amcHtml;

        // ====================================================================
        // TAB 4: SAMSUNG (Group by Category)
        // ====================================================================
        const samProdGrp = {};
        const samsungAllowedCats = ['AC', 'MICROWAVE OVEN', 'REFRIGERATOR', 'WASHING MACHINE', 'TV'];
        const pSam = fProduct.filter(r => r.brand && r.brand.toUpperCase().includes('SAMSUNG') && r.product && samsungAllowedCats.includes(r.product.toUpperCase().trim()));
        pSam.forEach(r => {
            const k = (r.product || 'Unknown').toUpperCase().trim();
            if (!samProdGrp[k]) samProdGrp[k] = { name: k, pQ: 0, pR: 0, sQ: 0, sR: 0 };
            samProdGrp[k].pQ += r.qty || 0;
            samProdGrp[k].pR += r.soldPrice || 0;
        });
        fSamsung.forEach(r => {
            const k = (r.product || 'Unknown').toUpperCase().trim();
            if (!samProdGrp[k]) samProdGrp[k] = { name: k, pQ: 0, pR: 0, sQ: 0, sR: 0 };
            samProdGrp[k].sQ += r.qty || 0;
            samProdGrp[k].sR += r.soldPrice || 0;
        });

        let samHtml = `<tr><th>Samsung Category</th><th>Prod Qty</th><th>Samsung Care Qty</th><th>Prod Rev</th><th>Samsung Care Rev</th><th>Qty Conv %</th><th>Val Conv %</th></tr>`;
        Object.values(samProdGrp).sort((a, b) => b.pQ - a.pQ).forEach(grp => {
            const qConv = grp.pQ > 0 ? ((grp.sQ / grp.pQ) * 100).toFixed(2) : '0.00';
            const vConv = grp.pR > 0 ? ((grp.sR / grp.pR) * 100).toFixed(2) : '0.00';
            samHtml += `<tr>
                <td><strong>${grp.name}</strong></td>
                <td class="col-num">${grp.pQ}</td>
                <td class="col-num">${grp.sQ}</td>
                <td class="col-num">${fmtShortHtml(grp.pR)}</td>
                <td class="col-num">${fmtShortHtml(grp.sR)}</td>
                <td class="col-num">${qConv}%</td>
                <td class="col-num">${vConv}%</td>
            </tr>`;
        });
        const tableSam = $('tableFsDashSamsung');
        if (tableSam && tableSam.querySelector('tbody')) tableSam.querySelector('tbody').innerHTML = samHtml;
    }

    // Main IIFE continues...

    // ---- CUSTOM CALENDAR WIDGET ----
window.coCalWidget = {
    _createDOM() {
        if (document.getElementById('coCustomCal')) return;
        const div = document.createElement('div');
        div.id = 'coCustomCal';
        div.style.cssText = "position:absolute; display:none; flex-direction:column; background:#1f2937; border:1px solid rgba(255,255,255,0.1); border-radius:12px; padding:16px; width:280px; box-shadow:0 10px 40px rgba(0,0,0,0.5); z-index:9999; font-family:'Inter', sans-serif;";

        div.innerHTML = `
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:16px;">
                <button id="coCalPrev" style="background:#111827; border:none; color:#fff; width:28px; height:28px; border-radius:8px; cursor:pointer;">&lt;</button>
                <div id="coCalMonth" style="color:#fff; font-weight:600; font-size:0.95rem;">April 2026</div>
                <button id="coCalNext" style="background:#111827; border:none; color:#fff; width:28px; height:28px; border-radius:8px; cursor:pointer;">&gt;</button>
            </div>
            <div style="display:grid; grid-template-columns:repeat(7, 1fr); gap:4px; text-align:center; color:#9ca3af; font-size:0.75rem; font-weight:600; margin-bottom:8px;">
                <div>Su</div><div>Mo</div><div>Tu</div><div>We</div><div>Th</div><div>Fr</div><div>Sa</div>
            </div>
            <div id="coCalDays" style="display:grid; grid-template-columns:repeat(7, 1fr); gap:4px; text-align:center;"></div>
            <div id="coCalFooter" style="margin-top:16px; border-top:1px solid rgba(255,255,255,0.05); padding-top:12px; font-size:0.75rem; color:#6b7280; text-align:center;">
                Click a date to set filter
            </div>
            <div style="margin-top:8px; text-align:center;">
                 <button id="coCalClear" style="background:rgba(239,68,68,0.1); border:1px solid rgba(239,68,68,0.3); color:#ef4444; padding:6px 14px; border-radius:6px; cursor:pointer; font-size:0.75rem; font-weight:600;">Clear Date</button>
            </div>
        `;
        document.body.appendChild(div);

        document.getElementById('coCalPrev').onclick = () => this.changeMonth(-1);
        document.getElementById('coCalNext').onclick = () => this.changeMonth(1);
        document.getElementById('coCalClear').onclick = () => {
            this.activeCallback('');
            this.close();
        };

        document.addEventListener('click', (e) => {
            if (this.isOpen && !div.contains(e.target) && e.target !== this.activeBtn && !this.activeBtn.contains(e.target)) {
                this.close();
            }
        });
    },
    isOpen: false,
    currentDate: new Date(),
    activeValue: '',
    activeCallback: null,
    activeBtn: null,

    _showFilter(btn) {
        let current_val = document.getElementById('coDateFilter') ? document.getElementById('coDateFilter').value : '';
        this.open(btn, current_val, (selected) => {
            if (document.getElementById('coDateFilter')) document.getElementById('coDateFilter').value = selected;
            if (document.getElementById('coDateFilterLabel')) document.getElementById('coDateFilterLabel').textContent = selected ? new Date(selected).toLocaleDateString('en-GB') : 'All Dates';
            if (typeof renderCustomersOSGPage === 'function') renderCustomersOSGPage();
        }, 'Click a date to filter invoices');
    },

    open(btn, initialVal, callback, footerText = 'Click a date to set follow-up') {
        this._createDOM();
        this.activeBtn = btn;
        this.activeCallback = callback;
        this.activeValue = initialVal;

        if (initialVal) this.currentDate = new Date(initialVal);
        else this.currentDate = new Date();

        document.getElementById('coCalFooter').textContent = footerText;

        const rect = btn.getBoundingClientRect();
        const cal = document.getElementById('coCustomCal');
        cal.style.top = (rect.bottom + window.scrollY + 8) + 'px';
        cal.style.left = (rect.left + window.scrollX) + 'px';
        cal.style.display = 'flex';
        this.isOpen = true;
        this.render();
    },
    close() {
        if (!this.isOpen) return;
        document.getElementById('coCustomCal').style.display = 'none';
        this.isOpen = false;
    },
    changeMonth(dir) {
        this.currentDate.setMonth(this.currentDate.getMonth() + dir);
        this.render();
    },
    render() {
        const y = this.currentDate.getFullYear();
        const m = this.currentDate.getMonth();
        document.getElementById('coCalMonth').textContent = new Date(y, m, 1).toLocaleString('default', { month: 'long', year: 'numeric' });

        const firstDay = new Date(y, m, 1).getDay();
        const daysInMon = new Date(y, m + 1, 0).getDate();

        let html = '';
        for (let i = 0; i < firstDay; i++) {
            html += `<div></div>`;
        }
        for (let d = 1; d <= daysInMon; d++) {
            const dateStr = `${y}-${String(m + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
            const isSelected = (dateStr === this.activeValue);

            let st = 'cursor:pointer; padding:6px 0; border-radius:6px; transition:all 0.1s; font-size:0.85rem; color:#d1d5db;';
            if (isSelected) {
                st += 'border:2px solid #f59e0b; color:#fff; font-weight:700;';
            } else {
                st += 'border:2px solid transparent;';
            }

            html += `<div class="cal-day" data-date="${dateStr}" style="${st}" onmouseover="this.style.background='rgba(255,255,255,0.05)'" onmouseout="this.style.background='transparent'">${d}</div>`;
        }
        document.getElementById('coCalDays').innerHTML = html;

        document.querySelectorAll('.cal-day').forEach(el => {
            el.onclick = (e) => {
                const val = e.target.getAttribute('data-date');
                // Toggle: clicking the already-selected date clears the filter
                const newVal = (val === this.activeValue) ? '' : val;
                this.activeValue = newVal;
                this.render();
                if (this.activeCallback) this.activeCallback(newVal);
                if (newVal) this.close();
            };
        });
    }
};

// ============================================================
// AI Business Agent Integration (Nova AI)
// ============================================================
document.addEventListener('DOMContentLoaded', function initAIAssistant() {
    const aiChatBtn = document.getElementById('aiChatBtn');
    const aiChatWindow = document.getElementById('aiChatWindow');
    const aiCloseBtn = document.getElementById('aiCloseBtn');
    const aiSettingsBtn = document.getElementById('aiSettingsBtn');
    const aiSettingsModal = document.getElementById('aiSettingsModal');
    const closeAiSettingsModal = document.getElementById('closeAiSettingsModal');
    const aiApiKeyInput = document.getElementById('aiApiKeyInput');
    const saveAiKeyBtn = document.getElementById('saveAiKeyBtn');
    const aiChatInput = document.getElementById('aiChatInput');
    const aiChatSendBtn = document.getElementById('aiChatSendBtn');
    const aiChatMessages = document.getElementById('aiChatMessages');

    if (!aiChatBtn) return; // Guard: exit if AI UI not in DOM

    // Load stored API Key
    var storedKey = localStorage.getItem('nova_ai_api_key') || '';
    if (storedKey && aiApiKeyInput) aiApiKeyInput.value = storedKey;

    // Toggle chat window
    aiChatBtn.addEventListener('click', function() {
        aiChatWindow.classList.toggle('hidden');
        if (!aiChatWindow.classList.contains('hidden') && aiChatInput) aiChatInput.focus();
    });
    aiCloseBtn.addEventListener('click', function() { aiChatWindow.classList.add('hidden'); });

    // Settings modal
    aiSettingsBtn.addEventListener('click', function() { aiSettingsModal.style.display = 'flex'; });
    closeAiSettingsModal.addEventListener('click', function() { aiSettingsModal.style.display = 'none'; });
    saveAiKeyBtn.addEventListener('click', function() {
        var key = aiApiKeyInput.value.trim();
        if (key) {
            localStorage.setItem('nova_ai_api_key', key);
            alert('API Key saved!');
            aiSettingsModal.style.display = 'none';
        } else {
            alert('Please enter a valid API key.');
        }
    });
    // ============================================================
    // ---- TARGET & FORECAST ENGINE ----
    // ============================================================
    (function initForecast() {
        const $ = id => document.getElementById(id);
        let fcDonutChartInstance = null;

        function renderForecastPage() {
            try {
                // Use allData to check if any data is loaded, as productData might be empty if only AMC was loaded
                if (typeof allData === 'undefined' || allData.length === 0) {
                    const kg = document.getElementById('fcKpiGrid');
                    if (kg) kg.style.display = 'none';
                    const nd = document.getElementById('fcNoData');
                    if (nd) nd.style.display = 'block';
                    return;
                }
                const nd = document.getElementById('fcNoData');
                if (nd) nd.style.display = 'none';

                const targetConv = parseFloat($('fcTargetConv').value) || 60;
                const totalDays  = parseInt($('fcTotalDays').value) || 26;
                const daysPassed = Math.min(parseInt($('fcDaysPassed').value) || 0, totalDays);
                const daysLeft   = Math.max(totalDays - daysPassed, 0);

                // Use filtered data if available, otherwise fall back to full dataset
                const activeProd = (typeof filteredProduct !== 'undefined' && filteredProduct.length > 0) ? filteredProduct : (typeof productData !== 'undefined' ? productData : []);
                const activeOSG  = (typeof filteredOSG !== 'undefined' && filteredOSG.length > 0) ? filteredOSG : (typeof osgData !== 'undefined' ? osgData : []);

                // Current stats
                const totalPQty = activeProd.reduce((s, r) => s + (r.qty || 0), 0);
                const totalOQty = activeOSG.reduce((s, r) => s + (r.qty || 0), 0);
                const currentConv = totalPQty > 0 ? (totalOQty / totalPQty) * 100 : 0;

                // How many OSG units are needed to hit the target?
                const targetOQty = Math.ceil((targetConv / 100) * totalPQty);
                const osgGap = Math.max(targetOQty - totalOQty, 0);
                const dailyRequired = daysLeft > 0 ? Math.ceil(osgGap / daysLeft) : 0;

                // Status colour
                const onTrack = currentConv >= targetConv;
                const statusColor = onTrack ? '#10b981' : (currentConv >= targetConv * 0.8 ? '#f59e0b' : '#ef4444');
                const statusLabel = onTrack ? 'ÃƒÂ°Ã…Â¸Ã…Â½Ã‚Â¯ On Target!' : (daysLeft === 0 ? 'ÃƒÂ¢Ã¢â‚¬ÂºÃ¢â‚¬Â Month Ended' : 'ÃƒÂ¢Ã…Â¡Ã‚Â ÃƒÂ¯Ã‚Â¸Ã‚Â Behind Target');

                // KPI cards
                const kpiDefs = [
                    { label: 'Current Conv %', value: currentConv.toFixed(1) + '%', color: statusColor },
                    { label: 'Target Conv %',  value: targetConv.toFixed(1) + '%',  color: '#6366f1' },
                    { label: 'OSG Gap',        value: osgGap + ' units',             color: '#ef4444' },
                    { label: 'Days Remaining', value: daysLeft + ' days',            color: '#3b82f6' },
                    { label: 'Daily Quota',    value: dailyRequired + ' OSG/day',    color: '#f97316' },
                    { label: 'Status',         value: statusLabel,                   color: statusColor },
                ];
                $('fcKpiCards').innerHTML = kpiDefs.map(k => `
                    <div style="background:var(--bg-card);border:1px solid var(--border);border-top:4px solid ${k.color};border-radius:12px;padding:18px 16px;">
                        <div style="font-size:0.78rem;font-weight:600;color:var(--text-muted);text-transform:uppercase;letter-spacing:0.5px;margin-bottom:8px;">${k.label}</div>
                        <div style="font-size:1.35rem;font-weight:800;color:${k.color};">${k.value}</div>
                    </div>`).join('');

                // Donut chart
                if (fcDonutChartInstance) { fcDonutChartInstance.destroy(); fcDonutChartInstance = null; }
                const donutCtx = document.getElementById('fcDonutChart');
                if (donutCtx && typeof Chart !== 'undefined') {
                    const achieved = Math.min(currentConv, targetConv);
                    const remaining = Math.max(targetConv - currentConv, 0);
                    const overshoot = Math.max(currentConv - targetConv, 0);
                    fcDonutChartInstance = new Chart(donutCtx.getContext('2d'), {
                        type: 'doughnut',
                        data: {
                            datasets: [{
                                data: onTrack ? [targetConv, overshoot, 100 - currentConv] : [achieved, remaining, 100 - targetConv],
                                backgroundColor: onTrack
                                    ? ['#10b981', '#6ee7b7', 'rgba(255,255,255,0.05)']
                                    : ['#6366f1', '#ef4444', 'rgba(255,255,255,0.05)'],
                                borderWidth: 0,
                                hoverOffset: 4
                            }]
                        },
                        options: {
                            cutout: '75%', responsive: true, maintainAspectRatio: false,
                            plugins: { legend: { display: false }, tooltip: { enabled: false } }
                        }
                    });
                    $('fcDonutLabel').innerHTML = `
                        <div style="font-size:1.6rem;font-weight:900;color:${statusColor};">${currentConv.toFixed(1)}%</div>
                        <div style="font-size:0.7rem;color:var(--text-muted);font-weight:600;">of ${targetConv}% target</div>`;
                }

                // Daily quota detail panel
                $('fcQuotaDetails').innerHTML = `
                    <div style="margin-bottom:14px;">
                        <div style="display:flex;justify-content:space-between;margin-bottom:6px;">
                            <span style="font-size:0.85rem;color:var(--text-muted);">OSG Sold so far</span>
                            <span style="font-weight:700;color:var(--text-primary);">${totalOQty.toLocaleString('en-IN')}</span>
                        </div>
                        <div style="display:flex;justify-content:space-between;margin-bottom:6px;">
                            <span style="font-size:0.85rem;color:var(--text-muted);">OSG Needed for ${targetConv}%</span>
                            <span style="font-weight:700;color:#6366f1;">${targetOQty.toLocaleString('en-IN')}</span>
                        </div>
                        <div style="display:flex;justify-content:space-between;margin-bottom:16px;">
                            <span style="font-size:0.85rem;color:var(--text-muted);">OSG Gap Remaining</span>
                            <span style="font-weight:700;color:#ef4444;">${osgGap.toLocaleString('en-IN')}</span>
                        </div>
                    </div>
                    <div style="background:rgba(99,102,241,0.08);border:1px solid rgba(99,102,241,0.2);border-radius:10px;padding:14px;text-align:center;">
                        <div style="font-size:0.75rem;color:#6366f1;font-weight:700;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:4px;">Daily Target</div>
                        <div style="font-size:2.2rem;font-weight:900;color:${daysLeft > 0 ? '#f97316' : '#ef4444'};">${daysLeft > 0 ? dailyRequired : 'N/A'}</div>
                        <div style="font-size:0.8rem;color:var(--text-muted);">OSG units/day over ${daysLeft} remaining days</div>
                    </div>
                    <div style="margin-top:14px;">
                        <div style="display:flex;justify-content:space-between;margin-bottom:4px;">
                            <span style="font-size:0.8rem;color:var(--text-muted);">Monthly Progress</span>
                            <span style="font-size:0.8rem;font-weight:700;color:var(--text-primary);">${Math.round(Math.min(currentConv/targetConv,1)*100)}%</span>
                        </div>
                        <div style="background:rgba(255,255,255,0.07);border-radius:8px;height:8px;overflow:hidden;">
                            <div style="height:100%;border-radius:8px;background:${statusColor};width:${Math.min(currentConv/targetConv*100,100).toFixed(1)}%;transition:width 0.6s ease;"></div>
                        </div>
                    </div>`;

                // RBM breakdown table
                const rbmMap = {};
                if (activeProd && activeProd.length) {
                    activeProd.forEach(r => {
                        if (!r.rbm) return;
                        if (!rbmMap[r.rbm]) rbmMap[r.rbm] = { pQty: 0, oQty: 0 };
                        rbmMap[r.rbm].pQty += (r.qty || 0);
                    });
                }
                if (activeOSG && activeOSG.length) {
                    activeOSG.forEach(r => {
                        const match = (typeof productData !== 'undefined' ? productData : []).find(p => p.invoice === r.invoice);
                        const rbm = match ? match.rbm : null;
                        if (rbm && rbmMap[rbm]) rbmMap[rbm].oQty += (r.qty || 0);
                    });
                }

                const rbmRows = Object.keys(rbmMap).map(rbm => {
                    const { pQty, oQty } = rbmMap[rbm];
                    const conv = pQty > 0 ? (oQty / pQty) * 100 : 0;
                    const rbmTargetOQty = Math.ceil((targetConv / 100) * pQty);
                    const rbmGap = Math.max(rbmTargetOQty - oQty, 0);
                    const rbmDaily = daysLeft > 0 ? Math.ceil(rbmGap / daysLeft) : 0;
                    const isOnTrack = conv >= targetConv;
                    return { rbm, pQty, oQty, conv, rbmGap, rbmDaily, isOnTrack };
                }).sort((a, b) => b.conv - a.conv);

                $('fcRbmTable').innerHTML = `
                    <div class="table-wrapper" style="overflow:auto;">
                        <table class="data-table" style="min-width:600px;">
                            <thead><tr>
                                <th>RBM</th>
                                <th style="text-align:right;">Products Sold</th>
                                <th style="text-align:right;">OSG Sold</th>
                                <th style="text-align:right;">Conv %</th>
                                <th style="text-align:right;">OSG Gap</th>
                                <th style="text-align:right;">OSG/Day Needed</th>
                                <th style="text-align:center;">Status</th>
                            </tr></thead>
                            <tbody>
                                ${rbmRows.map((r, i) => `<tr style="background:${i%2===0?'transparent':'rgba(255,255,255,0.02)'};">
                                    <td style="font-weight:600;color:var(--text-primary);">${r.rbm}</td>
                                    <td style="text-align:right;">${r.pQty.toLocaleString('en-IN')}</td>
                                    <td style="text-align:right;">${r.oQty.toLocaleString('en-IN')}</td>
                                    <td style="text-align:right;font-weight:700;color:${r.isOnTrack?'#10b981':'#ef4444'};">${r.conv.toFixed(1)}%</td>
                                    <td style="text-align:right;color:${r.rbmGap>0?'#ef4444':'#10b981'};font-weight:600;">${r.rbmGap > 0 ? r.rbmGap.toLocaleString('en-IN') : 'ÃƒÂ¢Ã…â€œÃ¢â‚¬Å“ Done'}</td>
                                    <td style="text-align:right;font-weight:700;color:${r.rbmDaily>0?'#f97316':'#10b981'};">${r.rbmDaily > 0 ? r.rbmDaily : '-'}</td>
                                    <td style="text-align:center;"><span style="padding:3px 10px;border-radius:20px;font-size:0.75rem;font-weight:700;background:${r.isOnTrack?'rgba(16,185,129,0.12)':'rgba(239,68,68,0.12)'};color:${r.isOnTrack?'#10b981':'#ef4444'};">${r.isOnTrack?'On Track':'Behind'}</span></td>
                                </tr>`).join('')}
                            </tbody>
                        </table>
                    </div>`;

                $('fcKpiGrid').style.display = 'block';
            } catch (err) {
                console.error("Forecast Error:", err);
                const nd = document.getElementById('fcNoData');
                if (nd) {
                    nd.innerHTML = `<div style="color:var(--text-primary);padding:30px;background:rgba(239,68,68,0.1);border:1px solid #ef4444;border-radius:12px;text-align:left;">
                        <h3 style="color:#ef4444;margin:0 0 10px 0;">Runtime Error</h3>
                        <p style="margin:0;font-family:monospace;font-size:0.9rem;">${err.message}</p>
                        <pre style="margin-top:10px;font-size:0.75rem;color:var(--text-muted);white-space:pre-wrap;">${err.stack}</pre>
                    </div>`;
                    nd.style.display = 'block';
                }
                const kg = document.getElementById('fcKpiGrid');
                if (kg) kg.style.display = 'none';
            }
        }

        // Expose globally so onclick attribute in HTML can call it
        window.renderForecastPage = renderForecastPage;

        // Nav click â€”Â auto-render when switching to this page
        var fcNavBtn = document.querySelector('[data-section="forecast-section"]');
        if (fcNavBtn) fcNavBtn.addEventListener('click', function() { setTimeout(renderForecastPage, 100); });

        // Calculate button â€”Â use getElementById directly (most reliable)
        var fcCalcBtn = document.getElementById('btnForecastCalc');
        if (fcCalcBtn) fcCalcBtn.addEventListener('click', renderForecastPage);
    })();

    // ============================================================
    // ---- AI CHAT ----
    // ============================================================

    // Simple markdown parser for AI responses
    function parseMarkdown(text) {
        var html = text.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
        html = html.replace(/\*(.*?)\*/g, '<em>$1</em>');
        html = html.replace(/\n/g, '<br>');
        return html;
    }

    function addMessage(text, sender) {
        var div = document.createElement('div');
        div.className = 'ai-message ' + sender;
        if (sender === 'assistant') {
            div.innerHTML = parseMarkdown(text);
        } else {
            div.textContent = text;
        }
        aiChatMessages.appendChild(div);
        aiChatMessages.scrollTop = aiChatMessages.scrollHeight;
    }

    function showTypingIndicator() {
        var div = document.createElement('div');
        div.className = 'ai-typing-indicator';
        div.id = 'aiTypingIndicator';
        div.innerHTML = '<div class="ai-typing-dot"></div><div class="ai-typing-dot"></div><div class="ai-typing-dot"></div>';
        aiChatMessages.appendChild(div);
        aiChatMessages.scrollTop = aiChatMessages.scrollHeight;
    }

    function removeTypingIndicator() {
        var ind = document.getElementById('aiTypingIndicator');
        if (ind) ind.remove();
    }

    // Build business context from live DOM KPIs
    function buildContext() {
        var pQty = (document.getElementById('kpiQty') || {}).textContent || 'N/A';
        var pRev = (document.getElementById('kpiRevenue') || {}).textContent || 'N/A';
        var oQty = (document.getElementById('kpiOsgQty') || {}).textContent || 'N/A';
        var oRev = (document.getElementById('kpiOsgRevenue') || {}).textContent || 'N/A';
        var qConv = (document.getElementById('kpiQtyConv') || {}).textContent || 'N/A';
        var vConv = (document.getElementById('kpiValConv') || {}).textContent || 'N/A';
        var amcQty = (document.getElementById('kpiAmcTotal') || {}).textContent || 'N/A';
        var amcConv = (document.getElementById('kpiAmcConv') || {}).textContent || 'N/A';
        var samQty = (document.getElementById('kpiSamsungTotal') || {}).textContent || 'N/A';
        var samConv = (document.getElementById('kpiSamsungConv') || {}).textContent || 'N/A';

        var extendedContext = '';
        if (window.getGodModeContextData) {
            try {
                var fullData = window.getGodModeContextData();
                extendedContext = '\n\nDETAILED AGGREGATE DATA (Use this to answer specific questions about staff, branches, RBMs, BDMs, missing CRM counts, or categories):\n' + JSON.stringify(fullData);
            } catch (e) {
                console.error("Failed to build god mode context", e);
            }
        }

        return 'You are Nova AI, a retail business intelligence analyst for an Indian electronics retail analytics portal. ' +
            'The live dashboard data right now is: ' +
            'Total Products Sold: ' + pQty + ', ' +
            'Total Revenue: ' + pRev + ', ' +
            'OSG Qty Sold: ' + oQty + ', ' +
            'OSG Revenue: ' + oRev + ', ' +
            'Qty Conversion Rate: ' + qConv + ', ' +
            'Value Conversion Rate: ' + vConv + ', ' +
            'LG AMC Qty: ' + amcQty + ', ' +
            'LG AMC Conversion: ' + amcConv + ', ' +
            'Samsung Qty: ' + samQty + ', ' +
            'Samsung Conversion: ' + samConv + '. ' +
            extendedContext + '\n\n' +
            'Answer questions about this data clearly and concisely using markdown formatting. ' +
            'Provide actionable suggestions when asked about improvements or losses. ' +
            'Keep answers under 200 words unless detail is specifically requested.';
    }

    // Call NVIDIA NIM API with Fallback
    async function sendMessageToAI(userMessage) {
        var apiKey = localStorage.getItem('nova_ai_api_key');
        if (!apiKey) {
            addMessage('I need an OpenRouter API Key to work. Click the âš™ï¸ Settings icon above to add it.', 'error');
            return;
        }
        addMessage(userMessage, 'user');
        aiChatInput.value = '';
        showTypingIndicator();

        var systemContext = buildContext();
        
        // Deep Search integration
        var isDeepSearch = document.getElementById('aiDeepSearchToggle') && document.getElementById('aiDeepSearchToggle').checked;
        if (isDeepSearch) {
            var allStaffData = typeof window.portalStaffStats !== 'undefined' ? window.portalStaffStats : [];
            var rawData = {
                Full_Staff_Stats: allStaffData.map(s => ({
                    StaffName: s.name,
                    Branch: s.branch,
                    RBM: s.rbm,
                    BDM: s.bdm,
                    Total_Products_Sold: s.pQty,
                    Total_Sold_Price_Revenue: s.pRev,
                    OSG_Products_Sold: s.oQty,
                    OSG_Sold_Price_Revenue: s.oRev,
                    Qty_Conversion_Percent: s.qtyConv ? s.qtyConv.toFixed(1) : 0,
                    Value_Conversion_Percent: s.valConv ? s.valConv.toFixed(1) : 0
                })),
                missingCRM: typeof missedUnique !== 'undefined' ? missedUnique : [],
                crmLogs: typeof coStatusMap !== 'undefined' ? coStatusMap : {}
            };
            systemContext += '\n\nDEEP SEARCH RAW DATASET (USE THIS FOR EXACT INVOICES OR STAFF LOOKUPS):\n' + JSON.stringify(rawData);
        }

        var payload = {
            messages: [
                { role: 'system', content: systemContext },
                { role: 'user', content: userMessage }
            ],
            temperature: 0.2,
            max_tokens: 2000
        };

        // OpenRouter.ai - Single key, 60+ models, natively supports browser CORS
        // Model IDs verified live from OpenRouter API on 2026-05-16
        var models = [
            'openrouter/free',                            // Auto-selects any working free model
            'deepseek/deepseek-v4-flash:free',            // Confirmed live - great for analysis
            'google/gemma-4-27b-it:free',                 // Confirmed live - Google model
            'nvidia/nemotron-3-super-120b-a12b:free'      // Confirmed live - large reasoning model
        ];

        // Use direct OpenRouter endpoint (supports Access-Control-Allow-Origin: *)
        var endpoint = 'https://openrouter.ai/api/v1/chat/completions';
        var aiText = null;
        var lastErr = null;

        for (let i = 0; i < models.length; i++) {
            payload.model = models[i];
            var controller = new AbortController();
            var timeoutId = setTimeout(function() { controller.abort(); }, 45000); // 45s timeout to allow free models to process large context
            try {
                var response = await fetch(endpoint, {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                        'Authorization': 'Bearer ' + apiKey
                    },
                    body: JSON.stringify(payload),
                    signal: controller.signal
                });
                clearTimeout(timeoutId);
                
                if (response.ok) {
                    var data = await response.json();
                    if (data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content) {
                        aiText = data.choices[0].message.content;
                        // Successfully got response, break fallback loop
                        break;
                    }
                } else {
                    var errData = await response.json();
                    lastErr = (errData.error && errData.error.message) || response.statusText || response.status;
                    console.warn(`Model ${models[i]} failed: ${lastErr}. Falling back...`);
                }
            } catch (err) {
                if (typeof timeoutId !== 'undefined') clearTimeout(timeoutId);
                lastErr = err.name === 'AbortError' ? 'Request timed out after 15 seconds' : err.message;
                console.warn(`Network error with ${models[i]}: ${lastErr}. Falling back...`);
            }
        }

        removeTypingIndicator();
        
        if (aiText) {
            addMessage(aiText, 'assistant');
        } else {
            addMessage('API Error across all models: ' + (lastErr || 'Unknown error'), 'error');
        }
    }

    aiChatSendBtn.addEventListener('click', function() {
        var msg = aiChatInput.value.trim();
        if (msg) sendMessageToAI(msg);
    });

    aiChatInput.addEventListener('keypress', function(e) {
        if (e.key === 'Enter') {
            var msg = aiChatInput.value.trim();
            if (msg) sendMessageToAI(msg);
        }
    });

});

// End of Main IIFE
})();
