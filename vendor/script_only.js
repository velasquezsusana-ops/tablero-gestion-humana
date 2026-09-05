
(function(){
"use strict";

/* ---------------------------------------------------------------- *
 *  Constants
 * ---------------------------------------------------------------- */
const MESES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
const TRIMESTRES = ["Trimestre 1","Trimestre 2","Trimestre 3","Trimestre 4"];
const SHEETS = {
  clima: "Clima organizacional",
  formacion: "Plan de formación",
  costo: "Costo de capacitación",
  contratacion: "Tiempo de contratación",
  accidentes: "Accidentes laborales"
};

function css(name){ return getComputedStyle(document.body).getPropertyValue(name).trim(); }
function mesIdx(m){ return MESES.indexOf(String(m||"").trim()); }
function periodKey(anio, mes){ return Number(anio)*100 + mesIdx(mes); }
function fmtPct(x){ return isFinite(x) ? (x*100).toLocaleString('es-CO', {maximumFractionDigits:1}) + '%' : '—'; }
function fmtMoney(x){ return isFinite(x) ? x.toLocaleString('es-CO', {style:'currency', currency:'COP', maximumFractionDigits:0}) : '—'; }
function fmtNum(x, d){ return isFinite(x) ? x.toLocaleString('es-CO', {maximumFractionDigits: d==null?0:d}) : '—'; }

/* ---------------------------------------------------------------- *
 *  IndexedDB — remember the picked file handle across sessions
 * ---------------------------------------------------------------- */
const IDB_NAME = "tablero-gh", IDB_STORE = "handles";
function idbOpen(){
  return new Promise((resolve, reject)=>{
    const req = indexedDB.open(IDB_NAME, 1);
    req.onupgradeneeded = ()=> req.result.createObjectStore(IDB_STORE);
    req.onsuccess = ()=> resolve(req.result);
    req.onerror = ()=> reject(req.error);
  });
}
async function idbSet(key, val){
  const db = await idbOpen();
  return new Promise((resolve, reject)=>{
    const tx = db.transaction(IDB_STORE, "readwrite");
    tx.objectStore(IDB_STORE).put(val, key);
    tx.oncomplete = ()=> resolve();
    tx.onerror = ()=> reject(tx.error);
  });
}
async function idbGet(key){
  const db = await idbOpen();
  return new Promise((resolve, reject)=>{
    const tx = db.transaction(IDB_STORE, "readonly");
    const req = tx.objectStore(IDB_STORE).get(key);
    req.onsuccess = ()=> resolve(req.result);
    req.onerror = ()=> reject(req.error);
  });
}

/* ---------------------------------------------------------------- *
 *  File connection: File System Access API with <input> fallback
 * ---------------------------------------------------------------- */
const supportsFSA = "showOpenFilePicker" in window;
let fileHandle = null;      // FileSystemFileHandle, when supported
let lastModified = 0;
let pollTimer = null;

const els = {
  fileLabel: document.getElementById('fileLabel'),
  statusDot: document.getElementById('statusDot'),
  statusText: document.getElementById('statusText'),
  btnConnect: document.getElementById('btnConnect'),
  btnConnect2: document.getElementById('btnConnect2'),
  btnRefresh: document.getElementById('btnRefresh'),
  emptyState: document.getElementById('emptyState'),
  views: document.getElementById('views'),
};

function setStatus(state, text){
  els.statusDot.className = 'dot ' + (state==='live'?'live':state==='off'?'off':'');
  els.statusText.textContent = text;
}

async function connectViaPicker(){
  if (supportsFSA){
    try{
      const [handle] = await window.showOpenFilePicker({
        types: [{ description: 'Excel', accept: {'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx']} }],
        excludeAcceptAllOption: false, multiple: false
      });
      fileHandle = handle;
      await idbSet('lastHandle', handle);
      await loadFromHandle();
      startPolling();
    }catch(err){ if (err.name !== 'AbortError') console.error(err); }
  } else {
    const input = document.createElement('input');
    input.type = 'file'; input.accept = '.xlsx';
    input.onchange = async ()=>{
      const f = input.files[0];
      if (!f) return;
      els.fileLabel.textContent = f.name + ' (selección manual — pulsa "Actualizar ahora" tras guardar cambios)';
      await loadWorkbookFromFile(f);
    };
    input.click();
  }
}

async function tryRestoreHandle(){
  if (!supportsFSA) return;
  try{
    const handle = await idbGet('lastHandle');
    if (!handle) return;
    const perm = await handle.queryPermission({mode:'read'});
    if (perm === 'granted'){
      fileHandle = handle;
      await loadFromHandle();
      startPolling();
    } else {
      // Needs a user gesture to (re)grant permission after a browser restart.
      fileHandle = handle;
      els.fileLabel.textContent = 'Archivo recordado — pulsa "Conectar Excel" para reactivar el acceso';
      setStatus('off', 'Reconexión requerida');
    }
  }catch(err){ /* handle may be stale (file moved/deleted) — ignore */ }
}

async function loadFromHandle(){
  const perm = await fileHandle.queryPermission({mode:'read'});
  if (perm !== 'granted'){
    const req = await fileHandle.requestPermission({mode:'read'});
    if (req !== 'granted') { setStatus('off','Permiso denegado'); return; }
  }
  const file = await fileHandle.getFile();
  lastModified = file.lastModified;
  els.fileLabel.textContent = file.name + ' — conectado';
  await loadWorkbookFromFile(file);
}

function startPolling(){
  if (pollTimer) clearInterval(pollTimer);
  pollTimer = setInterval(async ()=>{
    if (!fileHandle) return;
    try{
      const perm = await fileHandle.queryPermission({mode:'read'});
      if (perm !== 'granted') return;
      const file = await fileHandle.getFile();
      if (file.lastModified !== lastModified){
        lastModified = file.lastModified;
        await loadWorkbookFromFile(file);
        flashUpdated();
      }
    }catch(e){ /* file may be temporarily locked while Excel saves */ }
  }, 4000);
}

function flashUpdated(){
  const now = new Date();
  setStatus('live', 'Actualizado ' + now.toLocaleTimeString('es-CO', {hour:'2-digit', minute:'2-digit', second:'2-digit'}));
}

els.btnConnect.addEventListener('click', connectViaPicker);
els.btnConnect2.addEventListener('click', connectViaPicker);
els.btnRefresh.addEventListener('click', async ()=>{
  if (fileHandle) await loadFromHandle();
});

/* ---------------------------------------------------------------- *
 *  Workbook parsing
 * ---------------------------------------------------------------- */
let DATA = null; // parsed dataset

async function loadWorkbookFromFile(file){
  setStatus('', 'Leyendo archivo…');
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, {type:'array', cellDates:false});
  DATA = extractData(wb);
  els.emptyState.style.display = 'none';
  els.views.style.display = '';
  els.btnRefresh.disabled = false;
  renderAll();
  flashUpdated();
}

function findSheet(wb, wantedName){
  const name = wb.SheetNames.find(n => n.trim() === wantedName.trim());
  return name ? wb.Sheets[name] : null;
}

function sheetRows(ws, opts){
  if (!ws) return [];
  return XLSX.utils.sheet_to_json(ws, Object.assign({header:1, raw:true, defval: null}, opts||{}));
}

function extractData(wb){
  const out = { clima: [], formacion: [], costo: [], contratacionDetalle: [], contratacion: [], accidentes: [] };

  // Clima organizacional: B..I -> Año, Trimestre, Área, Respondientes, Promedio pts, %área, Objetivo(pts), Objetivo%
  const wsClima = findSheet(wb, SHEETS.clima);
  sheetRows(wsClima).slice(1).forEach(r=>{
    if (r[1]==null || r[3]==null) return;
    out.clima.push({ anio:r[1], trimestre:r[2], area:r[3], respondientes:r[4], promedio:r[5], pct:r[6], objetivoPts:r[7], objetivoPct:r[8] });
  });

  // Plan de formación: A..F -> Año, Mes, Planificadas, Realizadas, Cumplimiento, Objetivo
  const wsForm = findSheet(wb, SHEETS.formacion);
  sheetRows(wsForm).slice(1).forEach(r=>{
    if (r[0]==null || r[1]==null) return;
    out.formacion.push({ anio:r[0], mes:r[1], planificadas:r[2], realizadas:r[3], cumplimiento:r[4], objetivo:r[5] });
  });

  // Costo de capacitación: A..H -> Año, Mes, Papelería, Cursos, Exámenes, Total, Empleados, CostoPorEmpleado
  const wsCosto = findSheet(wb, SHEETS.costo);
  sheetRows(wsCosto).slice(1).forEach(r=>{
    if (r[0]==null || r[1]==null) return;
    out.costo.push({ anio:r[0], mes:r[1], papeleria:r[2], cursos:r[3], examenes:r[4], total:r[5], empleados:r[6], costoEmpleado:r[7] });
  });

  // Tiempo de contratación: A..F detail (Año, Mes, Vacante, Fecha apertura, Fecha ingreso, Días); I..M summary (Año, Mes, VacantesCubiertas, PromedioDías, Objetivo)
  const wsContr = findSheet(wb, SHEETS.contratacion);
  sheetRows(wsContr).slice(2).forEach(r=>{
    if (r[0]!=null && r[1]!=null && r[2]!=null){
      out.contratacionDetalle.push({ anio:r[0], mes:r[1], vacante:r[2], dias:r[5] });
    }
    if (r[8]!=null && r[9]!=null){
      out.contratacion.push({ anio:r[8], mes:r[9], vacantes:r[10], promedioDias:r[11], objetivo:r[12] });
    }
  });

  // Accidentes laborales: A..E -> Año, Mes, Empleados, Accidentes, Índice
  const wsAcc = findSheet(wb, SHEETS.accidentes);
  sheetRows(wsAcc).slice(1).forEach(r=>{
    if (r[0]==null || r[1]==null) return;
    out.accidentes.push({ anio:r[0], mes:r[1], empleados:r[2], accidentes:r[3], indice:r[4] });
  });

  const byPeriod = (a,b)=> periodKey(a.anio,a.mes) - periodKey(b.anio,b.mes);
  out.formacion.sort(byPeriod);
  out.costo.sort(byPeriod);
  out.contratacion.sort(byPeriod);
  out.accidentes.sort(byPeriod);
  out.clima.sort((a,b)=> (Number(a.anio)-Number(b.anio)) || (TRIMESTRES.indexOf(a.trimestre)-TRIMESTRES.indexOf(b.trimestre)));

  return out;
}

/* ---------------------------------------------------------------- *
 *  Tabs
 * ---------------------------------------------------------------- */
document.getElementById('tabs').addEventListener('click', (e)=>{
  const btn = e.target.closest('.tab');
  if (!btn) return;
  document.querySelectorAll('.tab').forEach(b=>b.classList.toggle('active', b===btn));
  document.querySelectorAll('.view').forEach(v=>v.classList.remove('active'));
  document.getElementById('view-'+btn.dataset.view).classList.add('active');
});

/* ---------------------------------------------------------------- *
 *  Chart helpers
 * ---------------------------------------------------------------- */
const charts = {};
function destroy(id){ if (charts[id]) { charts[id].destroy(); delete charts[id]; } }

const baseGrid = ()=>({ color: css('--grid'), drawTicks:false });
const baseTicks = ()=>({ color: css('--text-secondary'), font:{size:11} });

function lineDataset(label, data, color, opts){
  return Object.assign({
    label, data, borderColor: color, backgroundColor: color,
    pointRadius: 3, pointHoverRadius: 5, borderWidth: 2, tension: 0.25, fill: false
  }, opts||{});
}
function dashedTarget(label, data){
  return { label, data, borderColor: css('--text-muted'), backgroundColor: css('--text-muted'),
    borderWidth: 2, borderDash: [5,4], pointRadius: 0, tension: 0, fill: false, type: 'line', order: 0 };
}

function makeLineChart(canvasId, labels, datasets, yFmt){
  destroy(canvasId);
  const ctx = document.getElementById(canvasId);
  if (!ctx) return;
  charts[canvasId] = new Chart(ctx, {
    type: 'line',
    data: { labels, datasets },
    options: {
      responsive: true, maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { display: datasets.length > 1, labels: { color: css('--text-secondary'), boxWidth: 12, font:{size:11.5} } },
        tooltip: {
          backgroundColor: css('--surface-1'), titleColor: css('--text-primary'), bodyColor: css('--text-secondary'),
          borderColor: css('--border'), borderWidth: 1, padding: 10,
          callbacks: yFmt ? { label: (c)=> `${c.dataset.label}: ${yFmt(c.parsed.y)}` } : undefined
        }
      },
      scales: {
        x: { grid: { display:false }, ticks: baseTicks() },
        y: { grid: baseGrid(), ticks: Object.assign(baseTicks(), yFmt ? {callback:(v)=>yFmt(v)} : {}), beginAtZero: true }
      }
    }
  });
}

function makeBarChart(canvasId, labels, datasets, opts){
  destroy(canvasId);
  const ctx = document.getElementById(canvasId);
  if (!ctx) return;
  const yFmt = opts && opts.yFmt;
  charts[canvasId] = new Chart(ctx, {
    type: 'bar',
    data: { labels, datasets },
    options: {
      responsive: true, maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { display: datasets.length > 1, labels: { color: css('--text-secondary'), boxWidth: 12, font:{size:11.5} } },
        tooltip: {
          backgroundColor: css('--surface-1'), titleColor: css('--text-primary'), bodyColor: css('--text-secondary'),
          borderColor: css('--border'), borderWidth: 1, padding: 10,
          callbacks: yFmt ? { label: (c)=> `${c.dataset.label}: ${yFmt(c.parsed.y)}` } : undefined
        }
      },
      scales: {
        x: { stacked: !!(opts&&opts.stacked), grid: { display:false }, ticks: baseTicks() },
        y: { stacked: !!(opts&&opts.stacked), grid: baseGrid(), ticks: Object.assign(baseTicks(), yFmt?{callback:(v)=>yFmt(v)}:{}), beginAtZero: true }
      }
    }
  });
}

/* ---------------------------------------------------------------- *
 *  Render: filters population
 * ---------------------------------------------------------------- */
function fillSelect(sel, values, formatter){
  const cur = sel.value;
  sel.innerHTML = '';
  values.forEach(v=>{
    const opt = document.createElement('option');
    opt.value = v; opt.textContent = formatter ? formatter(v) : v;
    sel.appendChild(opt);
  });
  if (values.includes(cur)) sel.value = cur;
}

/* ---------------------------------------------------------------- *
 *  Render: master
 * ---------------------------------------------------------------- */
function renderAll(){
  if (!DATA) return;
  renderIndicadores();
  renderClima();
  renderFormacion();
  renderCosto();
  renderContratacion();
  renderAccidentes();
}

/* ---- INDICADORES (main) ---- */
function renderIndicadores(){
  const anios = [...new Set(DATA.clima.map(r=>r.anio))].sort((a,b)=>a-b);
  const fAnio = document.getElementById('fClimaAnio');
  const fTrim = document.getElementById('fClimaTrim');
  fillSelect(fAnio, anios);
  fillSelect(fTrim, TRIMESTRES);
  if (anios.length){
    const last = DATA.clima[DATA.clima.length-1];
    fAnio.value = last.anio; fTrim.value = last.trimestre;
  }
  fAnio.onchange = fTrim.onchange = renderClimaMainChart;
  renderClimaMainChart();

  // Plan de formación trend (all history)
  const fLabels = DATA.formacion.map(r=>`${r.mes.slice(0,3)} ${r.anio}`);
  makeLineChart('chartFormacionMain', fLabels, [
    lineDataset('Cumplimiento', DATA.formacion.map(r=>r.cumplimiento), css('--series-1')),
    dashedTarget('Objetivo', DATA.formacion.map(r=>r.objetivo))
  ], fmtPct);

  // Costo por empleado trend
  const cLabels = DATA.costo.map(r=>`${r.mes.slice(0,3)} ${r.anio}`);
  makeLineChart('chartCostoMain', cLabels, [
    lineDataset('Costo / empleado', DATA.costo.map(r=>r.costoEmpleado), css('--series-1'))
  ], fmtMoney);

  // Tiempo de contratación trend
  const tLabels = DATA.contratacion.map(r=>`${r.mes.slice(0,3)} ${r.anio}`);
  makeLineChart('chartContratacionMain', tLabels, [
    lineDataset('Promedio días', DATA.contratacion.map(r=>r.promedioDias), css('--series-1')),
    dashedTarget('Objetivo', DATA.contratacion.map(r=>r.objetivo))
  ], (v)=>fmtNum(v,1)+' d');

  // Accidentes trend
  const aLabels = DATA.accidentes.map(r=>`${r.mes.slice(0,3)} ${r.anio}`);
  makeLineChart('chartAccidentesMain', aLabels, [
    lineDataset('Índice de accidentes', DATA.accidentes.map(r=>r.indice), css('--series-1'))
  ], fmtPct);

  renderKpis();
}

function renderClimaMainChart(){
  const anio = Number(document.getElementById('fClimaAnio').value);
  const trim = document.getElementById('fClimaTrim').value;
  const rows = DATA.clima.filter(r=>Number(r.anio)===anio && r.trimestre===trim);
  const labels = rows.map(r=>r.area);
  makeBarChart('chartClimaMain', labels, [
    { label:'% área', data: rows.map(r=>r.pct), backgroundColor: css('--series-1'), borderRadius:4, maxBarThickness:42 },
    dashedTarget('Objetivo', rows.map(r=>r.objetivoPct))
  ], {yFmt: fmtPct});
  renderKpis();
}

function lastOf(arr){ return arr.length ? arr[arr.length-1] : null; }

function kpiCard(label, value, badgeText, badgeClass){
  const div = document.createElement('div');
  div.className = 'kpi';
  div.innerHTML = `<div class="label">${label}</div><div class="value">${value}</div>` +
    (badgeText ? `<div class="badge ${badgeClass}">${badgeClass==='good'?'✓':badgeClass==='critical'?'⚠':'•'} ${badgeText}</div>` : '');
  return div;
}

function renderKpis(){
  const grid = document.getElementById('kpiGrid');
  grid.innerHTML = '';

  // Clima (selected filter)
  const anio = document.getElementById('fClimaAnio').value;
  const trim = document.getElementById('fClimaTrim').value;
  const climaRows = DATA.clima.filter(r=>String(r.anio)===String(anio) && r.trimestre===trim);
  if (climaRows.length){
    const avgPct = climaRows.reduce((s,r)=>s+r.pct,0)/climaRows.length;
    const avgObj = climaRows.reduce((s,r)=>s+r.objetivoPct,0)/climaRows.length;
    const ok = avgPct >= avgObj;
    grid.appendChild(kpiCard('Clima organizacional (promedio)', fmtPct(avgPct),
      ok?'Cumple objetivo':'Bajo objetivo', ok?'good':'critical'));
  }

  const lf = lastOf(DATA.formacion);
  if (lf){
    const ok = lf.cumplimiento >= lf.objetivo;
    grid.appendChild(kpiCard(`Cumplimiento formación (${lf.mes} ${lf.anio})`, fmtPct(lf.cumplimiento),
      ok?'Cumple objetivo':'Bajo objetivo', ok?'good':'critical'));
  }

  const lc = lastOf(DATA.costo);
  if (lc){
    grid.appendChild(kpiCard(`Costo por empleado (${lc.mes} ${lc.anio})`, fmtMoney(lc.costoEmpleado), null, 'neutral'));
  }

  const lt = lastOf(DATA.contratacion);
  if (lt){
    const ok = lt.promedioDias <= lt.objetivo;
    grid.appendChild(kpiCard(`Tiempo de contratación (${lt.mes} ${lt.anio})`, fmtNum(lt.promedioDias,1)+' días',
      ok?'Cumple objetivo':'Sobre objetivo', ok?'good':'critical'));
  }

  const la = lastOf(DATA.accidentes);
  if (la){
    const ok = la.accidentes === 0;
    grid.appendChild(kpiCard(`Índice de accidentes (${la.mes} ${la.anio})`, fmtPct(la.indice),
      ok?'Sin accidentes':(la.accidentes+' accidente(s)'), ok?'good':(la.indice>0.02?'critical':'neutral')));
  }
}

/* ---- CLIMA tab ---- */
function renderClima(){
  const anios = [...new Set(DATA.clima.map(r=>r.anio))].sort((a,b)=>a-b);
  const sel = document.getElementById('climaAnioFilter');
  fillSelect(sel, ['Todos', ...anios]);
  sel.onchange = drawClimaTab;
  drawClimaTab();
}
function drawClimaTab(){
  const sel = document.getElementById('climaAnioFilter').value;
  const rows = sel==='Todos' ? DATA.clima : DATA.clima.filter(r=>String(r.anio)===String(sel));

  // grouped bar: x = Trimestre+Año, one dataset per Área
  const areas = [...new Set(rows.map(r=>r.area))];
  const periods = [...new Set(rows.map(r=>`${r.trimestre} ${r.anio}`))];
  const seriesColors = [css('--series-1'), css('--series-2'), css('--series-3'), css('--series-4'), css('--series-5')];
  const datasets = areas.map((area,i)=>({
    label: area,
    data: periods.map(p=>{
      const row = rows.find(r=>`${r.trimestre} ${r.anio}`===p && r.area===area);
      return row ? row.pct : null;
    }),
    backgroundColor: seriesColors[i % seriesColors.length], borderRadius:3, maxBarThickness:22
  }));
  makeBarChart('chartClimaFull', periods, datasets, {yFmt: fmtPct});

  // respondientes trend (sum by period)
  const respByPeriod = periods.map(p=> rows.filter(r=>`${r.trimestre} ${r.anio}`===p).reduce((s,r)=>s+(r.respondientes||0),0));
  makeBarChart('chartClimaResp', periods, [{ label:'Respondientes', data: respByPeriod, backgroundColor: css('--series-1'), borderRadius:4, maxBarThickness:34 }]);

  renderTable('tblClima',
    ['Año','Trimestre','Área','Respondientes','Promedio pts','% área','Objetivo %'],
    rows.map(r=>[r.anio, r.trimestre, r.area, fmtNum(r.respondientes), fmtNum(r.promedio,1), fmtPct(r.pct), fmtPct(r.objetivoPct)])
  );
}

/* ---- FORMACION tab ---- */
function renderFormacion(){
  const anios = [...new Set(DATA.formacion.map(r=>r.anio))].sort((a,b)=>a-b);
  const sel = document.getElementById('formacionAnioFilter');
  fillSelect(sel, ['Todos', ...anios]);
  sel.onchange = drawFormacionTab;
  drawFormacionTab();
}
function drawFormacionTab(){
  const sel = document.getElementById('formacionAnioFilter').value;
  const rows = sel==='Todos' ? DATA.formacion : DATA.formacion.filter(r=>String(r.anio)===String(sel));
  const labels = rows.map(r=>`${r.mes.slice(0,3)} ${r.anio}`);

  makeBarChart('chartFormacionCant', labels, [
    { label:'Planificadas', data: rows.map(r=>r.planificadas), backgroundColor: css('--series-1'), borderRadius:3, maxBarThickness:18 },
    { label:'Realizadas', data: rows.map(r=>r.realizadas), backgroundColor: css('--series-2'), borderRadius:3, maxBarThickness:18 }
  ]);

  makeLineChart('chartFormacionCumpl', labels, [
    lineDataset('Cumplimiento', rows.map(r=>r.cumplimiento), css('--series-1')),
    dashedTarget('Objetivo', rows.map(r=>r.objetivo))
  ], fmtPct);

  renderTable('tblFormacion',
    ['Año','Mes','Planificadas','Realizadas','Cumplimiento','Objetivo'],
    rows.map(r=>[r.anio, r.mes, fmtNum(r.planificadas), fmtNum(r.realizadas), fmtPct(r.cumplimiento), fmtPct(r.objetivo)])
  );
}

/* ---- COSTO tab ---- */
function renderCosto(){
  const anios = [...new Set(DATA.costo.map(r=>r.anio))].sort((a,b)=>a-b);
  const sel = document.getElementById('costoAnioFilter');
  fillSelect(sel, ['Todos', ...anios]);
  sel.onchange = drawCostoTab;
  drawCostoTab();
}
function drawCostoTab(){
  const sel = document.getElementById('costoAnioFilter').value;
  const rows = sel==='Todos' ? DATA.costo : DATA.costo.filter(r=>String(r.anio)===String(sel));
  const labels = rows.map(r=>`${r.mes.slice(0,3)} ${r.anio}`);

  makeBarChart('chartCostoDesglose', labels, [
    { label:'Papelería', data: rows.map(r=>r.papeleria), backgroundColor: css('--series-1'), borderRadius:2, maxBarThickness:22 },
    { label:'Cursos', data: rows.map(r=>r.cursos), backgroundColor: css('--series-2'), borderRadius:2, maxBarThickness:22 },
    { label:'Exámenes', data: rows.map(r=>r.examenes), backgroundColor: css('--series-3'), borderRadius:2, maxBarThickness:22 }
  ], {stacked:true, yFmt: fmtMoney});

  makeLineChart('chartCostoEmpleado', labels, [
    lineDataset('Costo / empleado', rows.map(r=>r.costoEmpleado), css('--series-1'))
  ], fmtMoney);

  renderTable('tblCosto',
    ['Año','Mes','Papelería','Cursos','Exámenes','Total','Empleados','Costo/empleado'],
    rows.map(r=>[r.anio, r.mes, fmtMoney(r.papeleria), fmtMoney(r.cursos), fmtMoney(r.examenes), fmtMoney(r.total), fmtNum(r.empleados), fmtMoney(r.costoEmpleado)])
  );
}

/* ---- CONTRATACION tab ---- */
function renderContratacion(){
  const anios = [...new Set(DATA.contratacion.map(r=>r.anio))].sort((a,b)=>a-b);
  const sel = document.getElementById('contratacionAnioFilter');
  fillSelect(sel, ['Todos', ...anios]);
  sel.onchange = drawContratacionTab;
  drawContratacionTab();
}
function drawContratacionTab(){
  const sel = document.getElementById('contratacionAnioFilter').value;
  const rows = sel==='Todos' ? DATA.contratacion : DATA.contratacion.filter(r=>String(r.anio)===String(sel));
  const labels = rows.map(r=>`${r.mes.slice(0,3)} ${r.anio}`);

  makeBarChart('chartContratacionVac', labels, [
    { label:'Vacantes cubiertas', data: rows.map(r=>r.vacantes), backgroundColor: css('--series-1'), borderRadius:4, maxBarThickness:24 }
  ]);

  makeLineChart('chartContratacionDias', labels, [
    lineDataset('Promedio días', rows.map(r=>r.promedioDias), css('--series-1')),
    dashedTarget('Objetivo', rows.map(r=>r.objetivo))
  ], (v)=>fmtNum(v,1)+' d');

  renderTable('tblContratacion',
    ['Año','Mes','Vacantes cubiertas','Promedio días','Objetivo'],
    rows.map(r=>[r.anio, r.mes, fmtNum(r.vacantes), fmtNum(r.promedioDias,1), fmtNum(r.objetivo,1)])
  );

  const detalle = sel==='Todos' ? DATA.contratacionDetalle : DATA.contratacionDetalle.filter(r=>String(r.anio)===String(sel));
  renderTable('tblContratacionDetalle',
    ['Año','Mes','Vacante','Días'],
    detalle.map(r=>[r.anio, r.mes, r.vacante, fmtNum(r.dias)])
  );
}

/* ---- ACCIDENTES tab ---- */
function renderAccidentes(){
  const anios = [...new Set(DATA.accidentes.map(r=>r.anio))].sort((a,b)=>a-b);
  const sel = document.getElementById('accidentesAnioFilter');
  fillSelect(sel, ['Todos', ...anios]);
  sel.onchange = drawAccidentesTab;
  drawAccidentesTab();
}
function drawAccidentesTab(){
  const sel = document.getElementById('accidentesAnioFilter').value;
  const rows = sel==='Todos' ? DATA.accidentes : DATA.accidentes.filter(r=>String(r.anio)===String(sel));
  const labels = rows.map(r=>`${r.mes.slice(0,3)} ${r.anio}`);

  makeBarChart('chartAccidentesCant', labels, [
    { label:'# Accidentes', data: rows.map(r=>r.accidentes), backgroundColor: css('--series-1'), borderRadius:4, maxBarThickness:24 }
  ]);

  makeLineChart('chartAccidentesIndice', labels, [
    lineDataset('Índice de accidentes', rows.map(r=>r.indice), css('--series-1'))
  ], fmtPct);

  renderTable('tblAccidentes',
    ['Año','Mes','# Empleados','# Accidentes','Índice'],
    rows.map(r=>[r.anio, r.mes, fmtNum(r.empleados), fmtNum(r.accidentes), fmtPct(r.indice)])
  );
}

/* ---- generic table renderer ---- */
function renderTable(tableId, headers, rows){
  const table = document.getElementById(tableId);
  const thead = '<thead><tr>' + headers.map(h=>`<th>${h}</th>`).join('') + '</tr></thead>';
  const tbody = '<tbody>' + rows.map(r=>'<tr>'+r.map(c=>`<td>${c==null?'—':c}</td>`).join('')+'</tr>').join('') + '</tbody>';
  table.innerHTML = thead + tbody;
}

/* ---------------------------------------------------------------- *
 *  Boot
 * ---------------------------------------------------------------- */
tryRestoreHandle();

})();
