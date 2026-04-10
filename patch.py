import re
with open('app.js', 'r', encoding='utf-8') as f:
    text = f.read()

widget_code = """
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
             if(document.getElementById('coDateFilter')) document.getElementById('coDateFilter').value = selected;
             if(document.getElementById('coDateFilterLabel')) document.getElementById('coDateFilterLabel').textContent = selected ? new Date(selected).toLocaleDateString('en-GB') : 'All Dates';
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
            const dateStr = `${y}-${String(m+1).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
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
                this.activeValue = val;
                this.render(); 
                if (this.activeCallback) this.activeCallback(val);
                this.close();
            };
        });
    }
};
"""

if 'window.coCalWidget =' not in text:
    text += '\\n' + widget_code
    with open('app.js', 'w', encoding='utf-8') as f:
        f.write(text)
    print('Widget appended!')
else:
    print('Widget already exists')
