#!/usr/bin/env python3
"""Adds the Version 11.0 'Rolling Orders' tab to index.html (idempotent)."""
import sys, re
src, dst = sys.argv[1], sys.argv[2]
s = open(src, encoding='utf-8').read()
if 'id="view-rolling"' in s:
    print('already patched'); open(dst, 'w', encoding='utf-8').write(s); sys.exit(0)

def sub_once(s, anchor, new, after=False, count_check=True):
    assert s.count(anchor) >= 1, 'anchor missing: ' + anchor[:70]
    i = s.index(anchor)
    return s[:i+len(anchor)] + new + s[i+len(anchor):] if after else s[:i] + new + s[i:]

# 1) main tab button (right after Productivity Detail)
s = sub_once(s, '<button class="main-tab" data-target="view-productivity">Productivity Detail</button>',
  '\n      <!-- v11.0 — Rolling Orders tab, placed right after Productivity Detail -->\n      <button class="main-tab" data-target="view-rolling">Rolling Orders</button>', after=True)

# 2) left-rail icon (after the Productivity Detail icon block)
rail_anchor = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 21V8l9-5 9 5v13"/><path d="M9 21v-6h6v6"/></svg>\n      </div>'
assert rail_anchor in s
s = sub_once(s, rail_anchor, '''
      <div class="rail-icon" title="Rolling Orders" data-view="rolling">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 10h18M9 4v16M15 4v16"/></svg>
      </div>''', after=True)

# 3) CSS
css = '''
  /* v11.0 — Rolling Orders tab */
  .rolling-ms{position:relative;}
  .rolling-ms-btn{font-family:'Inter';font-size:13px;padding:9px 12px;border:1px solid var(--line);border-radius:8px;background:var(--panel);min-width:230px;max-width:340px;text-align:left;cursor:pointer;display:flex;justify-content:space-between;gap:10px;align-items:center;}
  .rolling-ms-btn span.cv{color:var(--text-muted);font-size:11px;}
  .rolling-ms-panel{position:absolute;top:100%;left:0;margin-top:4px;z-index:30;background:var(--panel);border:1px solid var(--line);border-radius:10px;box-shadow:0 10px 28px rgba(16,24,38,.14);padding:8px;min-width:230px;max-height:340px;overflow:auto;}
  .rolling-ms-panel[hidden]{display:none;}
  .rolling-ms-panel label{display:flex;gap:9px;align-items:center;padding:6px 8px;border-radius:6px;font-size:13px;cursor:pointer;color:var(--text);text-transform:none;letter-spacing:0;font-weight:500;}
  .rolling-ms-panel label:hover{background:#f0f3f7;}
  .rolling-ms-panel hr{border:none;border-top:1px solid var(--line);margin:4px 0;}
  #rollingShot{background:var(--panel);padding:16px 18px 14px;border:1px solid var(--line);border-radius:12px;}
  .rolling-title{display:flex;justify-content:space-between;align-items:baseline;flex-wrap:wrap;gap:6px;margin-bottom:12px;}
  .rolling-title b{font-size:15px;color:var(--ink);} .rolling-title span{font-size:12px;color:var(--text-muted);}
  .rolling-scroll{overflow:auto;}
  table.rolling-table{border-collapse:separate;border-spacing:0;font-family:'Inter';font-size:12.5px;width:auto;min-width:720px;}
  .rolling-table tbody tr{cursor:default;}
  .rolling-table tbody td{font-family:'IBM Plex Mono',monospace;font-size:12.5px;}
  .rolling-table tbody tr:hover td{background-image:linear-gradient(rgba(16,24,38,.04),rgba(16,24,38,.04));}
  .rolling-table th,.rolling-table td{padding:8px 12px;text-align:right;border-bottom:1px solid var(--line);white-space:nowrap;font-variant-numeric:tabular-nums;}
  .rolling-table thead th{background:var(--ink);color:#fff;font-weight:600;position:sticky;top:0;font-size:11.5px;letter-spacing:.02em;text-transform:none;}
  .rolling-table thead th small{display:block;font-weight:500;opacity:.7;font-size:10.5px;margin-top:2px;}
  .rolling-table th.rl-corner{text-align:left;}
  .rolling-table tbody th{text-align:left;font-weight:600;color:var(--ink);background:#f6f8fa;position:sticky;left:0;}
  .rolling-table tbody th small{font-weight:500;color:var(--text-muted);margin-left:6px;}
  .rolling-table td.rl-na{color:#b6bfca;}
  .rolling-table tfoot th,.rolling-table tfoot td{font-weight:700;background:#eef2f6;color:var(--ink);border-top:2px solid var(--ink);border-bottom:none;}
  .rolling-table tfoot th{text-align:left;position:sticky;left:0;}
  #rollingNote{font-size:11.5px;color:var(--text-muted);margin-top:12px;line-height:1.55;}
  .rolling-empty{padding:34px;text-align:center;color:var(--text-muted);font-size:13px;}
'''
i = s.index('</style>')
s = s[:i] + css + s[i:]

# 4) view section (before Forecast view)
view = '''<div class="view" id="view-rolling">
        <div class="section-head">
          <div>
            <h2>Rolling Orders</h2>
            <div class="sub">Rolling order count by Order Date (columns) and Expected Delivery Date (rows) · refreshed 06:00 / 13:00 / 18:00 / 00:00 from the OIX Waybill History report</div>
          </div>
        </div>
        <div class="prod-controls">
          <div class="prod-field rolling-ms">
            <label for="rollingDistBtn">District</label>
            <button type="button" class="rolling-ms-btn" id="rollingDistBtn" aria-haspopup="true"><span id="rollingDistLabel">Overall</span><span class="cv">&#9662;</span></button>
            <div class="rolling-ms-panel" id="rollingDistPanel" hidden></div>
          </div>
          <div class="prod-field">
            <label for="rollingPctSelect">Percentage</label>
            <select id="rollingPctSelect">
              <option value="count">Count (no percentage)</option>
              <option value="pct">% of the order date's total</option>
            </select>
          </div>
          <div class="prod-field">
            <label for="rollingUnitSelect">Order / Waybill Count</label>
            <select id="rollingUnitSelect">
              <option value="order">Order count (distinct parent orders)</option>
              <option value="waybill">Waybill count</option>
            </select>
          </div>
        </div>
        <div id="rollingShot">
          <div class="rolling-title"><b id="rollingTitleMain">Rolling Order Trend</b><span id="rollingTitleSub"></span></div>
          <div class="rolling-scroll" id="rollingTableWrap"></div>
          <div id="rollingNote"></div>
        </div>
      </div>

      '''
s = sub_once(s, '<div class="view" id="view-forecast">', view)

# 5) data loading
load = '''    try{
      const resR = await fetchJSON('rolling_orders.json');
      if(!resR.ok) throw new Error('no rolling_orders.json');
      DATA.rolling = await resR.json();
    }catch(err){
      DATA.rolling = null; // v11.0 — tab shows an empty-state message until the first scheduled run
    }

'''
s = sub_once(s, '    computeOverviewDateBounds(); // v7.0 §1', load)

# 6) renderer + key map
s = s.replace('renderProductivityTab, renderForecastTab,', 'renderProductivityTab, renderRollingTab, renderForecastTab,', 1)
assert 'renderRollingTab, renderForecastTab' in s
s = s.replace("productivity:'view-productivity', forecast:'view-forecast',", "productivity:'view-productivity', rolling:'view-rolling', forecast:'view-forecast',", 1)
assert "rolling:'view-rolling'" in s

# 7) JS
js = r'''/* ---------------- v11.0 — Rolling Orders tab ----------------
   Data: rolling_orders.json (written by rolling_orders.py, 4x daily; also embedded in this file).
   days[orderDate] = { slot, asOf, cells: { deliveryDate: { OVERALL|district: [orders, waybills] } } }
   Rows = Expected Delivery Date, columns = Order Date (latest ROLLING_MAX_COLS). */
  const ROLLING_MAX_COLS = 14;
  const rollingState = { districts:new Set(['OVERALL']), pct:'count', unit:'order' };
  const ROLLING_WD = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
  function rollingFmtDate(iso){
    const [y,m,d] = iso.split('-').map(Number); const dt = new Date(Date.UTC(y,m-1,d));
    return { short: dt.toLocaleString('en-US',{month:'short',timeZone:'UTC'}) + '-' + String(d).padStart(2,'0'), wd: ROLLING_WD[dt.getUTCDay()] };
  }
  function rollingSlotLabel(slot){ return slot==='00:00' ? 'Final (00:00)' : (slot ? 'as of ' + slot : ''); }
  function rollingCellValue(cell){
    if(!cell) return null;
    const idx = rollingState.unit==='waybill' ? 1 : 0;
    if(rollingState.districts.has('OVERALL')) return (cell.OVERALL||[0,0])[idx];
    let sum = 0;
    rollingState.districts.forEach(d=>{ const v = cell[d]; if(v) sum += v[idx]; });
    return sum;
  }
  function rollingDistLabelText(){
    if(rollingState.districts.has('OVERALL')) return 'Overall';
    const arr = DISTRICTS.filter(d=>rollingState.districts.has(d));
    return arr.length > 3 ? arr.length + ' districts selected' : arr.join(', ');
  }
  function rollingBuildDistPanel(){
    const panel = document.getElementById('rollingDistPanel');
    if(!panel || panel.dataset.built) return;
    panel.dataset.built = '1';
    panel.innerHTML = '<label><input type="checkbox" value="OVERALL"> <b>Overall</b></label><hr>' +
      DISTRICTS.map(d=>`<label><input type="checkbox" value="${d}"> ${d}</label>`).join('');
  }
  function rollingSyncDistPanel(){
    document.querySelectorAll('#rollingDistPanel input').forEach(cb=>{ cb.checked = rollingState.districts.has(cb.value); });
    document.getElementById('rollingDistLabel').textContent = rollingDistLabelText();
  }
  function renderRollingTab(){
    rollingBuildDistPanel(); rollingSyncDistPanel();
    const wrap = document.getElementById('rollingTableWrap');
    const R = DATA.rolling;
    const unitTxt = rollingState.unit==='waybill' ? 'Waybill count' : 'Order count (parent orders)';
    const pctTxt = rollingState.pct==='pct' ? '% of order-date total' : 'count';
    document.getElementById('rollingTitleMain').textContent = 'Rolling Order Trend';
    if(!R || !R.days || !Object.keys(R.days).length){
      wrap.innerHTML = '<div class="rolling-empty">No rolling-order data yet — it appears after the first scheduled run of rolling_orders.py (06:00 / 13:00 / 18:00 / 00:00).</div>';
      document.getElementById('rollingTitleSub').textContent = '';
      document.getElementById('rollingNote').textContent = '';
      return;
    }
    const orderDates = Object.keys(R.days).sort().slice(-ROLLING_MAX_COLS);
    const delSet = new Set();
    orderDates.forEach(od=>Object.keys(R.days[od].cells||{}).forEach(k=>delSet.add(k)));
    const delDates = Array.from(delSet).sort();
    const colTotals = orderDates.map(od=>{
      let t = 0; let any = false;
      delDates.forEach(dd=>{ const v = rollingCellValue((R.days[od].cells||{})[dd]); if(v!=null){ t += v; any = true; } });
      return any ? t : null;
    });
    const colMax = orderDates.map((od,ci)=>{
      let m = 0; delDates.forEach(dd=>{ const v = rollingCellValue((R.days[od].cells||{})[dd]); if(v!=null && v>m) m = v; }); return m;
    });
    const fmt = v => v==null ? '–' : (rollingState.pct==='pct' ? v.toFixed(1) + '%' : v.toLocaleString('en-US'));
    let h = '<table class="rolling-table"><thead><tr><th class="rl-corner">Expected Delivery ↓ / Order Date →</th>';
    orderDates.forEach(od=>{ const f = rollingFmtDate(od); h += `<th>${f.short} (${f.wd})<small>${rollingSlotLabel(R.days[od].slot)}</small></th>`; });
    h += '</tr></thead><tbody>';
    delDates.forEach(dd=>{
      const f = rollingFmtDate(dd);
      h += `<tr><th>${f.short}<small>${f.wd}</small></th>`;
      orderDates.forEach((od,ci)=>{
        const raw = rollingCellValue((R.days[od].cells||{})[dd]);
        if(raw==null){ h += '<td class="rl-na">–</td>'; return; }
        const val = rollingState.pct==='pct' ? (colTotals[ci] ? raw/colTotals[ci]*100 : 0) : raw;
        const ratio = colMax[ci] ? raw/colMax[ci] : 0;
        h += `<td style="background:rgba(22,50,74,${(0.03+0.20*ratio).toFixed(3)})">${fmt(val)}</td>`;
      });
      h += '</tr>';
    });
    h += '</tbody><tfoot><tr><th>Total</th>';
    orderDates.forEach((od,ci)=>{
      const t = colTotals[ci];
      h += `<td>${t==null ? '–' : (rollingState.pct==='pct' ? '100.0%' : t.toLocaleString('en-US'))}</td>`;
    });
    h += '</tr></tfoot></table>';
    wrap.innerHTML = h;
    const last = orderDates[orderDates.length-1];
    document.getElementById('rollingTitleSub').textContent =
      `${rollingDistLabelText()} · ${unitTxt} · ${pctTxt} · latest order date ${rollingFmtDate(last).short} ${rollingSlotLabel(R.days[last].slot)}`;
    document.getElementById('rollingNote').innerHTML =
      'Order Date = date the order was created (cancelled waybills removed). Each column is overwritten by the latest session of that day; the 00:00 session is the final daily summary. ' +
      'Orders = distinct parent order numbers; when several districts are selected the order counts are the sum of each district\'s distinct parent orders. ' +
      '"Overall" covers every district including waybills that cannot be mapped to one of the 10 districts (e.g. Macau). ' +
      (R.generatedAt ? 'Data generated ' + R.generatedAt + '.' : '');
  }
  (function wireRolling(){
    const btn = document.getElementById('rollingDistBtn'), panel = document.getElementById('rollingDistPanel');
    btn.addEventListener('click', e=>{ e.stopPropagation(); rollingBuildDistPanel(); rollingSyncDistPanel(); panel.hidden = !panel.hidden; });
    document.addEventListener('click', e=>{ if(!panel.hidden && !e.target.closest('.rolling-ms')) panel.hidden = true; });
    panel.addEventListener('change', e=>{
      const cb = e.target; if(!cb || cb.tagName!=='INPUT') return;
      const st = rollingState.districts;
      if(cb.value==='OVERALL'){ st.clear(); st.add('OVERALL'); }
      else{
        st.delete('OVERALL');
        if(cb.checked) st.add(cb.value); else st.delete(cb.value);
        if(!st.size) st.add('OVERALL');
      }
      renderRollingTab();
    });
    document.getElementById('rollingPctSelect').addEventListener('change', e=>{ rollingState.pct = e.target.value; renderRollingTab(); });
    document.getElementById('rollingUnitSelect').addEventListener('change', e=>{ rollingState.unit = e.target.value; renderRollingTab(); });
  })();

  '''
s = sub_once(s, '/* ---------------- tab switching ---------------- */', js)
open(dst, 'w', encoding='utf-8').write(s)
print('patched OK', len(s))
