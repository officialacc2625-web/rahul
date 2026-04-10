import re

with open('app.js', 'r', encoding='utf-8') as f:
    text = f.read()

# Update buildCoRowHTML completely
old_build_fn = text[text.find('function buildCoRowHTML(r, i) {'):text.find('function exportCustomersOSGExcel() {')]

new_build_fn = """function buildCoRowHTML(r, i) {
        const inv = r.invoice || String(i);
        const st  = coStatusMap[inv] || { callStatus: null, interest: null, calledBy: null, remarks: '' };

        const rowBg = st.interest === 'interested'     ? 'rgba(22,163,74,0.06)'
                    : st.interest === 'not-interested' ? 'rgba(220,38,38,0.06)'
                    : 'transparent';
                    
        // Locking logic
        const isLocked = st.calledBy && st.calledBy !== currentCaller;
        const disAttr = isLocked ? 'disabled' : '';
        const opcStyle = isLocked ? 'opacity:0.5; cursor:not-allowed;' : '';

        // Added 'follow-up' and 'bought' options to Interest
        const interestBtnOptions = `
            <option value="" ${!st.interest ? 'selected' : ''}>- Select -</option>
            <option value="interested" ${st.interest === 'interested' ? 'selected' : ''}>✅ Interested</option>
            <option value="not-interested" ${st.interest === 'not-interested' ? 'selected' : ''}>❌ Not Interested</option>
            <option value="follow-up" ${st.interest === 'follow-up' ? 'selected' : ''}>📅 Follow-up</option>
            <option value="bought" ${st.interest === 'bought' ? 'selected' : ''}>🛒 Bought</option>
        `;

        let setDateBtn = '';
        if (st.interest === 'follow-up') {
            const fDate = st.followUpDate ? new Date(st.followUpDate).toLocaleDateString('en-GB') : 'Set date';
            setDateBtn = `<button ${disAttr} style="margin-top:6px; padding:6px 10px; border-radius:6px; border:1px solid rgba(255,255,255,0.1); background:#111827; color:#d1d5db; font-size:0.75rem; font-family:inherit; cursor:pointer; width:100%; display:flex; align-items:center; gap:6px; ${opcStyle}" onclick="window.coCalWidget.open(this, '${st.followUpDate||''}', (val) => window.setCoFollowUpDate('${inv}', val))"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg> ${fDate}</button>`;
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
            const c = CO_CALLERS.find(x => x.name === st.calledBy) || { color:'#64748b', bg:'rgba(100,116,139,0.15)' };
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
                    <option value="connected" ${st.callStatus === 'connected' ? 'selected' : ''}>📞 Connected</option>
                    <option value="not-connected" ${st.callStatus === 'not-connected' ? 'selected' : ''}>🔕 Not Connected</option>
                    <option value="disconnected" ${st.callStatus === 'disconnected' ? 'selected' : ''}>📵 Disconnected</option>
                </select>
                ${callerInfo}
            </div>` : '';

        const remarksInput = `
            <input ${disAttr} type="text" value="${q(st.remarks || '')}" onchange="window.saveCoRemark('${inv}', this.value)" placeholder="Remarks..." style="
                padding:6px 10px;border-radius:8px;border:1px solid var(--border);
                font-size:0.85rem;font-family:inherit;background:var(--bg-input);color:var(--text-primary);
                width:130px; outline:none; ${opcStyle}" />
        `;

        let dStr = '—';
        let dt = r.time || r.invoiceDate;
        if (dt) dStr = new Date(dt).toLocaleString('en-GB', { day:'numeric', month:'numeric', year:'numeric', hour:'numeric', minute:'numeric' });

        return `<tr id="co-row-${inv}" style="border-bottom:1px solid var(--border);background:${rowBg};transition:background 0.2s;">
            <td style="padding:12px 10px;color:var(--text-muted);font-size:0.8rem;">${i+1}</td>
            <td style="padding:12px 10px;font-family:monospace;font-size:0.82rem;color:var(--text-secondary); width:120px;">${dStr}</td>
            <td style="padding:12px 10px;"><strong style="color:var(--text-primary);font-size:0.9rem;">${r.customerName||'—'}</strong><div style="font-size:0.75rem;color:var(--text-muted)">${r.invoice}</div></td>
            <td style="padding:12px 10px; width:150px;">${interestBtn}</td>
            <td style="padding:12px 10px; width:160px;">
                <div style="display:flex;align-items:center;gap:6px;">
                    <span style="font-weight:600;color:var(--text-primary);font-size:0.88rem;">${r.customerNo||'—'}</span>
                    ${r.customerNo?`<a href="tel:${r.customerNo}" title="Call" style="color:var(--primary);display:flex;padding:5px;border-radius:50%;background:rgba(59,130,246,0.12);"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72 12.84 12.84 0 0 0 .7 2.81 2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45 12.84 12.84 0 0 0 2.81.7A2 2 0 0 1 22 16.92z"/></svg></a>`:''}
                    ${r.customerNo?`<a href="https://wa.me/91${r.customerNo.replace(/\D/g,'')}" target="_blank" title="WhatsApp" style="color:#25D366;display:flex;padding:5px;border-radius:50%;background:rgba(37,211,102,0.12);"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 11.5a8.38 8.38 0 0 1-.9 3.8 8.5 8.5 0 0 1-7.6 4.7 8.38 8.38 0 0 1-3.8-.9L3 21l1.9-5.7a8.38 8.38 0 0 1-.9-3.8 8.5 8.5 0 0 1 4.7-7.6 8.38 8.38 0 0 1 3.8-.9h.5a8.48 8.48 0 0 1 8 8v.5z"/></svg></a>`:''}
                </div>
                ${callBtns}
            </td>
            <td style="padding:12px 10px;">${remarksInput}</td>
            <td style="padding:12px 10px;color:var(--text-secondary);font-size:0.85rem;">${r.branch||'—'}</td>
            <td style="padding:12px 10px;color:var(--text-secondary);font-size:0.85rem;">${r.product||'—'}</td>
            <td style="padding:12px 10px;text-align:right;font-weight:600;color:var(--text-primary);font-size:0.88rem;white-space:nowrap;">${fmtShort(r.soldPrice)}</td>
        </tr>`;
    }
"""

if old_build_fn:
    text = text.replace(old_build_fn, new_build_fn)
    with open('app.js', 'w', encoding='utf-8') as f:
        f.write(text)
    print("Replaced buildCoRowHTML successfully")
else:
    print("Failed to find old function")
