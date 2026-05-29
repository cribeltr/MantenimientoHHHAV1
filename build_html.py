#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Genera un archivo HTML autocontenido con dos vistas interactivas:
  1) Registro_MP-2026  (columnas B..AQ, filas 8..973)
  2) PMP_2026          (Carta Gantt)
Ambas vistas comparten las mismas funcionalidades: buscador global,
filtros por columna tipo Excel (listas desplegables con casillas),
ordenamiento, contador de filas, leyenda y exportación a CSV.
"""
import json
import os
import sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))

# Origen del libro Excel: 1) argumento CLI  2) copia en el repo  3) ruta de subida
CANDIDATES = [
    sys.argv[1] if len(sys.argv) > 1 else None,
    os.path.join(HERE, "Programacion_MP_2026.xlsm"),
    "/root/.claude/uploads/758a1371-4358-481e-b55d-b82f2a18dc47/73e4b702-Programacio_nMP_2026.xlsm",
]
SRC = next((p for p in CANDIDATES if p and os.path.exists(p)), None)
if not SRC:
    raise SystemExit("No se encontró el archivo Excel. Pase la ruta como argumento.")
OUT = os.path.join(HERE, "programacion_mp_2026.html")

wb = openpyxl.load_workbook(SRC, data_only=True)


def norm(v):
    """Normaliza un valor de celda a texto limpio para el HTML."""
    if v is None:
        return ""
    if isinstance(v, float):
        if v.is_integer():
            return str(int(v))
        return str(v)
    if isinstance(v, str):
        return v.strip()
    return str(v)


# ----------------------------------------------------------------------------
# Definición de columnas
# ----------------------------------------------------------------------------
MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun",
         "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]

# ---- Registro_MP-2026 : columnas B(2) .. AQ(43) ----
reg_cols = []
# Info: B..S  (col 2..19)
reg_info_idx = list(range(2, 20))
for c in reg_info_idx:
    label = norm(wb["Registro_MP-2026"].cell(row=7, column=c).value)
    reg_cols.append({"col": c, "label": label, "group": None, "type": "text"})
# Meses: T..AQ (col 20..43), pares P/R
c = 20
for m in MESES:
    reg_cols.append({"col": c,     "label": "P", "group": m, "type": "status"})
    reg_cols.append({"col": c + 1, "label": "R", "group": m, "type": "status"})
    c += 2

# ---- PMP_2026 : A(1)..AE(31) + AG(33),AH(34),AI(35)  (se omite AF vacía) ----
pmp_cols = []
pmp_info_idx = list(range(1, 20))  # A..S
for c in pmp_info_idx:
    label = norm(wb["PMP_2026"].cell(row=7, column=c).value)
    pmp_cols.append({"col": c, "label": label, "group": None, "type": "text"})
# Meses T..AE (20..31)
c = 20
for m in MESES:
    pmp_cols.append({"col": c, "label": m, "group": None, "type": "status"})
    c += 1
# Resumen AG,AH,AI (33,34,35)
for c in (33, 34, 35):
    label = norm(wb["PMP_2026"].cell(row=7, column=c).value)
    pmp_cols.append({"col": c, "label": label, "group": None, "type": "num"})


def extract(sheet, cols, r0=8, r1=973):
    ws = wb[sheet]
    rows = []
    for r in range(r0, r1 + 1):
        # Solo filas con ID (columna B = 2) no vacío
        if norm(ws.cell(row=r, column=2).value) == "":
            continue
        rows.append([norm(ws.cell(row=r, column=cd["col"]).value) for cd in cols])
    return rows


reg_rows = extract("Registro_MP-2026", reg_cols)
pmp_rows = extract("PMP_2026", pmp_cols)

# Quitamos la clave interna "col" antes de exportar
for cd in reg_cols:
    cd.pop("col", None)
for cd in pmp_cols:
    cd.pop("col", None)

views = {
    "registro": {
        "id": "registro",
        "title": "Registro MP 2026",
        "subtitle": "Registro de mantenciones preventivas realizadas — Hospital Hernán Henríquez Aravena",
        "columns": reg_cols,
        "rows": reg_rows,
        "stickyCols": 2,   # ID + N° Carpeta
        "labelCol": 3,     # índice de "Equipo" para mostrar en chips/filtros (0-based)
    },
    "pmp": {
        "id": "pmp",
        "title": "PMP 2026 · Carta Gantt",
        "subtitle": "Programa de mantención preventiva (carta Gantt) — Hospital Hernán Henríquez Aravena",
        "columns": pmp_cols,
        "rows": pmp_rows,
        "stickyCols": 2,
        "labelCol": 4,
    },
}

data_json = json.dumps(views, ensure_ascii=False)
print("Registro filas:", len(reg_rows), "columnas:", len(reg_cols))
print("PMP filas:", len(pmp_rows), "columnas:", len(pmp_cols))

# ----------------------------------------------------------------------------
# Plantilla HTML  (los datos se inyectan en __DATA__)
# ----------------------------------------------------------------------------
HTML = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Programación MP 2026 · HHHA</title>
<style>
  :root{
    --bg:#0f172a; --panel:#ffffff; --ink:#1e293b; --muted:#64748b;
    --line:#e2e8f0; --head:#1e3a5f; --head2:#27496d; --accent:#2563eb;
    --hover:#eff6ff; --chip:#e2e8f0;
  }
  *{box-sizing:border-box}
  body{margin:0;font-family:"Segoe UI",Roboto,Helvetica,Arial,sans-serif;
       color:var(--ink);background:#f1f5f9;font-size:13px}
  header.top{background:linear-gradient(135deg,#1e3a5f,#2563eb);color:#fff;
       padding:14px 20px;box-shadow:0 2px 8px rgba(0,0,0,.15)}
  header.top h1{margin:0;font-size:18px;font-weight:600}
  header.top p{margin:3px 0 0;font-size:12px;opacity:.85}

  .tabs{display:flex;gap:4px;background:#e2e8f0;padding:8px 20px 0}
  .tab{padding:9px 18px;border:none;border-radius:8px 8px 0 0;cursor:pointer;
       background:#cbd5e1;color:#334155;font-size:13px;font-weight:600}
  .tab.active{background:var(--panel);color:var(--accent);box-shadow:0 -2px 4px rgba(0,0,0,.05)}

  .toolbar{display:flex;flex-wrap:wrap;gap:10px;align-items:center;
       padding:10px 20px;background:var(--panel);border-bottom:1px solid var(--line)}
  .toolbar input[type=search]{padding:7px 11px;border:1px solid #cbd5e1;border-radius:8px;
       width:280px;font-size:13px}
  .btn{padding:7px 13px;border:1px solid #cbd5e1;background:#fff;border-radius:8px;
       cursor:pointer;font-size:12px;font-weight:600;color:#334155}
  .btn:hover{background:#f1f5f9}
  .btn.primary{background:var(--accent);color:#fff;border-color:var(--accent)}
  .btn.primary:hover{background:#1d4ed8}
  .count{margin-left:auto;font-size:12px;color:var(--muted);font-weight:600}
  .count b{color:var(--ink)}

  .legend{display:flex;flex-wrap:wrap;gap:7px 14px;padding:8px 20px;background:#f8fafc;
       border-bottom:1px solid var(--line);font-size:11px;color:var(--muted)}
  .legend .lg{display:inline-flex;align-items:center;gap:5px}
  .legend .sw{width:14px;height:14px;border-radius:3px;display:inline-block;border:1px solid rgba(0,0,0,.1)}

  .wrap{padding:0 20px 28px}
  .tablebox{overflow:auto;max-height:calc(100vh - 230px);border:1px solid var(--line);
       border-radius:8px;background:#fff;position:relative}
  table{border-collapse:separate;border-spacing:0;white-space:nowrap}
  th,td{border-right:1px solid var(--line);border-bottom:1px solid var(--line);
       padding:5px 8px;text-align:left;font-weight:400;vertical-align:middle}
  td .cell{max-width:230px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  td.mes,thead th.mes{min-width:30px;max-width:48px;text-align:center}
  thead th{position:sticky;top:0;background:var(--head);color:#fff;font-weight:600;
       z-index:3;font-size:12px}
  thead tr.r2 th{top:29px;background:var(--head2)}
  thead th.grp{text-align:center}
  th .htxt{display:inline-flex;align-items:center;gap:6px}
  .fbtn{border:none;background:rgba(255,255,255,.18);color:#fff;border-radius:4px;
       cursor:pointer;font-size:10px;line-height:1;padding:3px 5px}
  .fbtn:hover{background:rgba(255,255,255,.35)}
  .fbtn.on{background:#fbbf24;color:#1e293b}
  tbody tr:nth-child(even){background:#f8fafc}
  tbody tr:hover{background:var(--hover)}
  tbody td{font-size:12px}
  tbody tr.hidden{display:none}

  /* columnas fijas (sticky-left) */
  th.sticky,td.sticky{position:sticky;left:0;z-index:2}
  thead th.sticky{z-index:5;background:var(--head)}
  td.sticky{background:#fff}
  tbody tr:nth-child(even) td.sticky{background:#f1f5f9}
  tbody tr:hover td.sticky{background:var(--hover)}
  td.idx{color:var(--muted);font-variant-numeric:tabular-nums}

  /* estados / colores */
  td.st{text-align:center;font-weight:700;font-size:11px}
  .st-x   {background:#dbeafe;color:#1e40af}
  .st-r   {background:#dcfce7;color:#166534}
  .st-pm  {background:#ffedd5;color:#9a3412}
  .st-ra  {background:#ede9fe;color:#5b21b6}
  .st-si  {background:#bbf7d0;color:#14532d}
  .st-sira{background:#99f6e4;color:#115e59}
  .st-causa{background:#fee2e2;color:#991b1b}
  .st-fs  {background:#fef9c3;color:#854d0e}
  .st-baja{background:#e2e8f0;color:#475569}
  .st-no  {background:#fecaca;color:#7f1d1d}
  .st-nu  {background:#f1f5f9;color:#64748b}
  .sw-x{background:#dbeafe}.sw-r{background:#dcfce7}.sw-pm{background:#ffedd5}
  .sw-ra{background:#ede9fe}.sw-si{background:#bbf7d0}.sw-sira{background:#99f6e4}
  .sw-causa{background:#fee2e2}.sw-fs{background:#fef9c3}.sw-baja{background:#e2e8f0}
  .sw-no{background:#fecaca}.sw-nu{background:#f1f5f9}

  /* popup filtro */
  .popup{position:absolute;z-index:50;background:#fff;border:1px solid #cbd5e1;
       border-radius:10px;box-shadow:0 10px 30px rgba(0,0,0,.22);width:250px;
       padding:10px;display:none;font-size:12px;color:var(--ink)}
  .popup.show{display:block}
  .popup .srt{display:flex;gap:6px;margin-bottom:8px}
  .popup .srt .btn{flex:1;text-align:center;padding:5px}
  .popup input.fsearch{width:100%;padding:6px 8px;border:1px solid #cbd5e1;
       border-radius:6px;margin-bottom:8px}
  .popup .list{max-height:230px;overflow:auto;border:1px solid var(--line);
       border-radius:6px;padding:4px}
  .popup .opt{display:flex;align-items:center;gap:7px;padding:3px 5px;border-radius:4px;cursor:pointer}
  .popup .opt:hover{background:#f1f5f9}
  .popup .opt input{cursor:pointer}
  .popup .opt span{overflow:hidden;text-overflow:ellipsis}
  .popup .acts{display:flex;gap:6px;margin-top:9px}
  .popup .acts .btn{flex:1;text-align:center}
  .empty{padding:30px;text-align:center;color:var(--muted)}

  /* botón historial en la celda ID + selector de pendientes */
  .toolbar select{padding:7px 10px;border:1px solid #cbd5e1;border-radius:8px;font-size:12px;
       font-weight:600;color:#334155;background:#fff;cursor:pointer}
  .hbtn{border:none;background:#e0e7ff;color:#3730a3;border-radius:5px;cursor:pointer;
       font-size:11px;padding:1px 5px;margin-right:5px;vertical-align:middle;line-height:1.4}
  .hbtn:hover{background:#c7d2fe}
  .hbtn .hb{font-weight:700;margin-left:2px}
  .hbtn.has{background:#fde68a;color:#92400e}
  .hbtn.ovd{background:#fecaca;color:#991b1b}
  td.clickrow{cursor:pointer}

  /* estados pendiente */
  .st-abierto{background:#fef3c7;color:#92400e}
  .st-proceso{background:#dbeafe;color:#1e40af}
  .st-cerrado{background:#dcfce7;color:#166534}

  /* panel lateral (drawer) */
  .overlay{position:fixed;inset:0;background:rgba(15,23,42,.45);z-index:90;display:none}
  .overlay.show{display:block}
  .drawer{position:fixed;top:0;right:0;height:100vh;width:min(580px,97vw);background:#f1f5f9;
       z-index:100;box-shadow:-8px 0 30px rgba(0,0,0,.25);transform:translateX(102%);
       transition:transform .22s ease;display:flex;flex-direction:column}
  .drawer.show{transform:none}
  .dhead{background:linear-gradient(135deg,#1e3a5f,#2563eb);color:#fff;padding:13px 18px;
       display:flex;justify-content:space-between;align-items:flex-start;gap:10px}
  .dhead h3{margin:0;font-size:16px}
  .dsub{font-size:12px;opacity:.92;margin-top:4px;line-height:1.55}
  .dclose{background:rgba(255,255,255,.2);border:none;color:#fff;border-radius:6px;
       cursor:pointer;font-size:15px;padding:4px 11px;flex:none}
  .dbody{padding:14px 16px;overflow:auto;flex:1}
  .ev{background:#fff;border:1px solid var(--line);border-left:4px solid #2563eb;
       border-radius:8px;padding:10px 12px;margin-bottom:12px}
  .evh{display:flex;justify-content:space-between;align-items:center;gap:8px}
  .evdate{font-weight:700;color:#1e3a5f}
  .evdesc{margin:6px 0;color:#334155;white-space:pre-wrap}
  .pend{background:#f8fafc;border:1px solid var(--line);border-radius:7px;padding:8px 10px;margin:8px 0 0}
  .pend.ovd{border-color:#ef4444;background:#fef2f2}
  .ph{display:flex;justify-content:space-between;gap:8px;align-items:center}
  .pt{font-weight:600;color:#1e293b}
  .task{display:flex;align-items:flex-start;gap:7px;padding:5px 0;border-top:1px dashed var(--line)}
  .task.done .tdesc{text-decoration:line-through;color:#94a3b8}
  .task .tdesc{flex:1}
  .meta{font-size:11px;color:var(--muted);margin-top:2px}
  .chip{display:inline-block;padding:1px 8px;border-radius:10px;font-size:10px;font-weight:700;white-space:nowrap}
  .chip.abierto{background:#fef3c7;color:#92400e}
  .chip.proceso{background:#dbeafe;color:#1e40af}
  .chip.cerrado{background:#dcfce7;color:#166534}
  .chip.tarea{background:#e0e7ff;color:#3730a3}
  .chip.gestion{background:#fae8ff;color:#86198f}
  .chip.tipoev{background:#e2e8f0;color:#475569}
  .minibtn{border:none;background:transparent;cursor:pointer;color:#64748b;font-size:13px;
       padding:2px 5px;border-radius:4px}
  .minibtn:hover{background:#e2e8f0}
  .addbtn{border:1px dashed #94a3b8;background:#fff;color:#334155;border-radius:7px;cursor:pointer;
       font-size:12px;padding:6px 10px;width:100%;margin-top:7px;font-weight:600}
  .addbtn:hover{background:#eff6ff;border-color:#2563eb;color:#2563eb}
  .addbtn.sub{padding:4px 8px;font-size:11px}
  .form{background:#fff;border:1px solid #93c5fd;border-radius:8px;padding:10px;margin:8px 0}
  .form label{display:block;font-size:11px;color:#475569;margin:6px 0 2px;font-weight:600}
  .form input,.form select,.form textarea{width:100%;padding:6px 8px;border:1px solid #cbd5e1;
       border-radius:6px;font-size:12px;font-family:inherit;color:var(--ink)}
  .form .frow{display:flex;gap:8px}.form .frow>div{flex:1}
  .form .facts{display:flex;gap:7px;margin-top:9px}
  .form .facts .btn{flex:1;text-align:center}
  .dhint{font-size:11px;color:var(--muted);text-align:center;padding:6px 0 2px}
</style>
</head>
<body>
<header class="top">
  <h1>Programación de Mantención Preventiva 2026</h1>
  <p>Equipos médicos críticos · Hospital Hernán Henríquez Aravena</p>
</header>

<div class="tabs" id="tabs"></div>
<div class="toolbar">
  <input type="search" id="globalSearch" placeholder="🔍 Buscar en todas las columnas...">
  <select id="histFilter" title="Filtrar por estado de pendientes">
    <option value="all">📋 Pendientes: todos</option>
    <option value="open">Con pendientes abiertos</option>
    <option value="overdue">Con pendientes vencidos</option>
  </select>
  <button class="btn" id="clearFilters">✕ Limpiar filtros</button>
  <button class="btn primary" id="exportCsv">⬇ Exportar CSV</button>
  <button class="btn" id="expHist" title="Descargar respaldo del historial">⬇ Historial</button>
  <button class="btn" id="impHist" title="Cargar respaldo del historial">⬆ Importar</button>
  <input type="file" id="impFile" accept="application/json,.json" style="display:none">
  <span class="count" id="rowCount"></span>
</div>
<div class="legend" id="legend"></div>
<div class="wrap">
  <div class="tablebox" id="tablebox"></div>
</div>

<div class="popup" id="popup"></div>

<!-- Panel lateral: historial de eventos / pendientes / tareas por equipo -->
<div class="overlay" id="overlay"></div>
<aside class="drawer" id="drawer" aria-label="Historial del equipo">
  <div class="dhead">
    <div><h3 id="dTitle"></h3><div class="dsub" id="dSub"></div></div>
    <button class="dclose" id="dClose" title="Cerrar">✕</button>
  </div>
  <div class="dbody" id="dBody"></div>
</aside>

<script>
const DATA = __DATA__;

/* ---------- utilidades de estado/color ---------- */
function statusClass(v){
  if(!v) return "";
  const t = String(v).trim();
  if(/^C\d$/i.test(t)) return "st-causa";
  const m = {
    "X":"st-x","R":"st-r","PM":"st-pm","RA":"st-ra",
    "Si":"st-si","Sí":"st-si","Si-RA":"st-sira","Sí-RA":"st-sira",
    "FS":"st-fs","Baja":"st-baja","No":"st-no","NU":"st-nu",
    "abierto":"st-abierto","en proceso":"st-proceso","cerrado":"st-cerrado",
    "pendiente":"st-nu","hecha":"st-si"
  };
  return m[t] || "";
}
const LEGEND = [
  ["sw-x","X – Programado"],["sw-r","R – Reprogramado (P)"],
  ["sw-pm","PM – Pendiente mes"],["sw-ra","RA – Realizado anticipado"],
  ["sw-si","Si – Realizada"],["sw-sira","Si-RA – Realizada (reprog.)"],
  ["sw-causa","C1–C8 – Causa reprogramación"],["sw-fs","FS – Fuera de servicio"],
  ["sw-baja","Baja"],["sw-no","No realizada"],["sw-nu","NU – No utilizado"]
];

/* ---------- estado global ---------- */
let CURRENT = "registro";
const state = {};   // por vista: {filters:{idx:Set}, sort:{idx,dir}, search:""}
for(const k in DATA){ state[k] = {filters:{}, sort:null, search:"", histFilter:"all"}; }

const $ = s => document.querySelector(s);
const tabsEl = $("#tabs"), boxEl = $("#tablebox"), popup = $("#popup");

/* ---------- pestañas ---------- */
function buildTabs(){
  tabsEl.innerHTML = "";
  for(const k in DATA){
    const b = document.createElement("button");
    b.className = "tab" + (k===CURRENT?" active":"");
    b.textContent = DATA[k].title;
    b.onclick = () => { CURRENT = k; buildTabs(); buildLegend(); renderTable(); };
    tabsEl.appendChild(b);
  }
}

function buildLegend(){
  $("#legend").innerHTML = LEGEND.map(([c,t]) =>
    `<span class="lg"><span class="sw ${c}"></span>${t}</span>`).join("");
}

/* ---------- construcción de la tabla ---------- */
function renderTable(){
  const v = DATA[CURRENT];
  if(CURRENT === "pendientes") rebuildPendientes();
  syncToolbar();
  const cols = v.columns;
  const hasGroups = cols.some(c => c.group);
  const sticky = v.stickyCols || 0;
  const isEq = v.action !== false;   // vistas de equipos (Registro / PMP)

  // THEAD
  let thead = "<thead>";
  // fila 1
  thead += "<tr class='r1'>";
  let i = 0;
  while(i < cols.length){
    const c = cols[i];
    if(c.group){
      // agrupar columnas consecutivas con mismo group
      let j = i; while(j < cols.length && cols[j].group === c.group) j++;
      thead += `<th class='grp' colspan='${j-i}'>${esc(c.group)}</th>`;
      i = j;
    }else{
      let cl = (i < sticky ? "sticky " : "") + (c.type==="status" && isEq ? "mes" : "");
      thead += `<th class='${cl.trim()}' rowspan='${hasGroups?2:1}' style='${i<sticky?leftStyle(cols,i,sticky):""}'>${headInner(c,i)}</th>`;
      i++;
    }
  }
  thead += "</tr>";
  // fila 2 (solo subcolumnas de grupos)
  if(hasGroups){
    thead += "<tr class='r2'>";
    cols.forEach((c,idx) => { if(c.group) thead += `<th class='mes'>${headInner(c,idx)}</th>`; });
    thead += "</tr>";
  }
  thead += "</thead>";

  // TBODY
  const rowsHtml = v.rows.map((row,ri) => {
    let tds = "";
    cols.forEach((c,idx) => {
      const val = row[idx] ?? "";
      let cls = "", inner;
      if(c.type === "status"){
        const sc = statusClass(val); cls = "st" + (sc?(" "+sc):"") + (isEq?" mes":"");
        inner = esc(val);
      }else{
        inner = `<div class='cell' title="${esc(val)}">${esc(val)}</div>`;
      }
      // botón de historial en la columna ID (solo vistas de equipos)
      if(isEq && idx === v.idCol){ inner = histBtnHtml(val) + inner; }
      let style = "";
      if(idx < sticky){ cls += (cls?" ":"") + "sticky"; style = leftStyle(cols,idx,sticky); }
      tds += `<td class='${cls}' style='${style}'>${inner}</td>`;
    });
    // en la vista Pendientes la fila completa abre el historial del equipo
    const clk = (v.action===false) ? " class='clickrow'" : "";
    return `<tr data-ri='${ri}'${clk}>${tds}</tr>`;
  }).join("");

  boxEl.innerHTML = `<table>${thead}<tbody>${rowsHtml}</tbody></table>`;
  applyFilters();
}

function headInner(c, idx){
  const f = state[CURRENT].filters[idx];
  const on = f && f.size ? " on" : "";
  return `<span class='htxt'>${esc(c.label)}`+
         `<button class='fbtn${on}' data-idx='${idx}' title='Filtrar / ordenar'>▼</button></span>`;
}

// desplazamiento acumulado para columnas sticky
function leftStyle(cols, idx, sticky){
  // ancho aproximado por columna previa (se calcula tras render para exactitud)
  return "left:" + (window.__lefts && window.__lefts[idx]!=null ? window.__lefts[idx] : (idx*120)) + "px;";
}

function esc(s){ return String(s).replace(/[&<>"']/g, m =>
  ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m])); }

/* recalcula offsets exactos de columnas sticky tras render */
function recalcSticky(){
  const v = DATA[CURRENT]; const sticky = v.stickyCols||0;
  if(!sticky) return;
  const firstRow = boxEl.querySelector("tbody tr");
  if(!firstRow) return;
  const tds = firstRow.children;
  let acc = 0; const lefts = [];
  for(let i=0;i<sticky;i++){ lefts[i]=acc; acc += tds[i].offsetWidth; }
  window.__lefts = lefts;   // referencia para futuros renders
  // aplicar offsets exactos a header y cuerpo
  const heads = boxEl.querySelectorAll("thead th.sticky");
  heads.forEach((th,n)=>{ th.style.left = lefts[n] + "px"; });
  boxEl.querySelectorAll("tbody tr").forEach(tr=>{
    for(let i=0;i<sticky;i++){ if(tr.children[i]) tr.children[i].style.left = lefts[i]+"px"; }
  });
}

/* ---------- filtrado ---------- */
function applyFilters(){
  const v = DATA[CURRENT];
  const st = state[CURRENT];
  const search = st.search.trim().toLowerCase();
  const fEntries = Object.entries(st.filters).filter(([,s]) => s && s.size);

  const hf = st.histFilter || "all";
  const isEq = v.action !== false;

  const trs = boxEl.querySelectorAll("tbody tr");
  let shown = 0;
  trs.forEach(tr => {
    const row = v.rows[+tr.dataset.ri];
    let ok = true;
    // filtros por columna
    for(const [idx,set] of fEntries){
      if(!set.has(row[idx] ?? "")){ ok = false; break; }
    }
    // filtro por estado de pendientes (solo vistas de equipos)
    if(ok && isEq && hf !== "all"){
      const ps = pendStats(String(row[v.idCol]));
      if(hf === "open" && ps.open <= 0) ok = false;
      else if(hf === "overdue" && !ps.overdue) ok = false;
    }
    // búsqueda global
    if(ok && search){
      ok = row.some(cell => String(cell).toLowerCase().includes(search));
    }
    tr.classList.toggle("hidden", !ok);
    if(ok) shown++;
  });
  $("#rowCount").innerHTML = `Mostrando <b>${shown}</b> de <b>${v.rows.length}</b> ${v.unit || "equipos"}`;
  recalcSticky();
}

/* ---------- ordenamiento ---------- */
function sortBy(idx, dir){
  const v = DATA[CURRENT];
  const num = v.columns[idx].type === "num";
  v.rows.sort((a,b)=>{
    let x=a[idx]??"", y=b[idx]??"";
    if(num){ x=parseFloat(x)||0; y=parseFloat(y)||0; return dir==="asc"? x-y : y-x; }
    x=String(x).toLowerCase(); y=String(y).toLowerCase();
    if(x<y) return dir==="asc"?-1:1;
    if(x>y) return dir==="asc"?1:-1;
    return 0;
  });
  state[CURRENT].sort = {idx,dir};
  renderTable();
}

/* ---------- popup de filtro (lista desplegable tipo Excel) ---------- */
let popupIdx = null;
function openPopup(idx, anchor){
  popupIdx = idx;
  const v = DATA[CURRENT];
  const st = state[CURRENT];
  // valores distintos de la columna
  const counts = new Map();
  v.rows.forEach(r => { const val = r[idx] ?? ""; counts.set(val, (counts.get(val)||0)+1); });
  const values = [...counts.keys()].sort((a,b)=>{
    const na=parseFloat(a), nb=parseFloat(b);
    if(!isNaN(na)&&!isNaN(nb)) return na-nb;
    return String(a).localeCompare(String(b),"es");
  });
  const sel = st.filters[idx] ? st.filters[idx] : new Set(values); // sin filtro = todos
  const noFilter = !st.filters[idx];

  let h = `<div class='srt'>
      <button class='btn' data-act='asc'>▲ A-Z</button>
      <button class='btn' data-act='desc'>▼ Z-A</button>
    </div>
    <input class='fsearch' placeholder='Buscar valor...'>
    <div class='list'>
      <label class='opt'><input type='checkbox' class='all' ${ (noFilter||sel.size===values.length)?'checked':'' }><b>(Seleccionar todo)</b></label>`;
  values.forEach(val=>{
    const lbl = val==="" ? "(vacías)" : esc(val);
    const ck = (noFilter || sel.has(val)) ? "checked" : "";
    h += `<label class='opt'><input type='checkbox' class='vchk' data-val='${esc(val)}' ${ck}><span title="${lbl}">${lbl}</span> <span style='color:#94a3b8'>(${counts.get(val)})</span></label>`;
  });
  h += `</div>
    <div class='acts'>
      <button class='btn' data-act='clear'>Limpiar</button>
      <button class='btn primary' data-act='apply'>Aceptar</button>
    </div>`;
  popup.innerHTML = h;
  popup.classList.add("show");

  // posición
  const r = anchor.getBoundingClientRect();
  const px = Math.min(r.left + window.scrollX, window.scrollX + document.documentElement.clientWidth - 270);
  popup.style.left = Math.max(8, px) + "px";
  popup.style.top  = (r.bottom + window.scrollY + 4) + "px";

  // eventos internos
  popup.querySelector(".fsearch").oninput = e=>{
    const q = e.target.value.toLowerCase();
    popup.querySelectorAll(".list .opt").forEach(o=>{
      if(o.querySelector(".all")) return;
      const t = o.textContent.toLowerCase();
      o.style.display = t.includes(q) ? "" : "none";
    });
  };
  popup.querySelector(".all").onchange = e=>{
    popup.querySelectorAll(".vchk").forEach(c=>{ if(c.closest(".opt").style.display!=="none") c.checked = e.target.checked; });
  };
  popup.querySelectorAll("[data-act]").forEach(b=>{
    b.onclick = ()=>{
      const act = b.dataset.act;
      if(act==="asc"||act==="desc"){ closePopup(); sortBy(idx,act); return; }
      if(act==="clear"){
        delete st.filters[idx];
        closePopup(); renderTable(); return;
      }
      if(act==="apply"){
        const chosen = new Set();
        popup.querySelectorAll(".vchk").forEach(c=>{ if(c.checked) chosen.add(c.dataset.val); });
        if(chosen.size===values.length){ delete st.filters[idx]; }
        else { st.filters[idx] = chosen; }
        closePopup(); renderTable(); return;
      }
    };
  });
}
function closePopup(){ popup.classList.remove("show"); popupIdx=null; }

/* ---------- eventos globales ---------- */
boxEl.addEventListener("click", e=>{
  const fb = e.target.closest(".fbtn");
  if(fb){ e.stopPropagation(); openPopup(+fb.dataset.idx, fb); return; }
  const hb = e.target.closest(".hbtn");
  if(hb){ e.stopPropagation(); openDrawer(hb.dataset.eq); return; }
  // en la vista Pendientes, clic en la fila abre el historial del equipo
  if(DATA[CURRENT].action === false){
    const tr = e.target.closest("tr[data-ri]");
    if(tr){ const eq = DATA.pendientes._eq[+tr.dataset.ri]; if(eq) openDrawer(eq); }
  }
});
document.addEventListener("click", e=>{
  if(popup.classList.contains("show") && !popup.contains(e.target) && !e.target.closest(".fbtn"))
    closePopup();
});
$("#globalSearch").addEventListener("input", e=>{
  state[CURRENT].search = e.target.value; applyFilters();
});
$("#clearFilters").addEventListener("click", ()=>{
  state[CURRENT].filters = {}; state[CURRENT].search = ""; state[CURRENT].histFilter = "all";
  $("#globalSearch").value = ""; renderTable();
});
$("#histFilter").addEventListener("change", e=>{
  state[CURRENT].histFilter = e.target.value; applyFilters();
});
$("#exportCsv").addEventListener("click", exportCsv);
window.addEventListener("resize", recalcSticky);

function exportCsv(){
  const v = DATA[CURRENT];
  const cols = v.columns;
  const head = cols.map(c => (c.group? c.group+" " : "") + c.label);
  const lines = [head];
  boxEl.querySelectorAll("tbody tr").forEach(tr=>{
    if(tr.classList.contains("hidden")) return;
    lines.push(v.rows[+tr.dataset.ri]);
  });
  const csv = lines.map(r => r.map(cell=>{
    const s = String(cell ?? "");
    return /[",\n;]/.test(s) ? '"'+s.replace(/"/g,'""')+'"' : s;
  }).join(";")).join("\r\n");
  const blob = new Blob(["﻿"+csv], {type:"text/csv;charset=utf-8"});
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  const names = {registro:"Registro_MP_2026", pmp:"PMP_2026", pendientes:"Pendientes_MP_2026"};
  a.download = (names[CURRENT] || "MP_2026") + "_filtrado.csv";
  a.click();
}

/* ===================================================================
   HISTORIAL POR EQUIPO  (eventos -> pendientes -> tareas/gestiones)
   Persistencia: localStorage + respaldo JSON. Indexado por ID de equipo.
   =================================================================== */
const HKEY = "mp2026_historial_v1";
let HIST = {};
let CURRENT_EQ = null;        // equipo abierto en el panel
let editing = null;           // formulario abierto: {kind,mode,evId,pId,tId}
const TIPOS_EVENTO = ["Falla","Mantención correctiva","Mantención preventiva",
                      "Inspección","Reprogramación","Observación","Otro"];
let REG_COL = {}, REG_BY_ID = {};

function todayStr(){ const d=new Date(); return new Date(d.getTime()-d.getTimezoneOffset()*60000).toISOString().slice(0,10); }
function uid(){ return "x"+Math.random().toString(36).slice(2,9)+Date.now().toString(36).slice(-4); }

function loadHist(){ try{ HIST = JSON.parse(localStorage.getItem(HKEY) || "{}") || {}; }catch(e){ HIST = {}; } }
function saveHist(){
  try{ localStorage.setItem(HKEY, JSON.stringify(HIST)); }
  catch(e){ alert("No se pudo guardar el historial (almacenamiento lleno o bloqueado). Usa «⬇ Historial» para respaldar."); }
  afterChange();
}
function afterChange(){
  refreshBadges();
  if(CURRENT === "pendientes"){ rebuildPendientes(); renderTable(); }
}
function getHist(id){ return HIST[id] || {eventos:[]}; }            // solo lectura
function ensureEq(id){ if(!HIST[id]) HIST[id] = {eventos:[]}; return HIST[id]; }  // mutable

function pendStats(id){
  const h = HIST[id]; let open=0, total=0, overdue=false;
  if(h) for(const ev of (h.eventos||[])) for(const p of (ev.pendientes||[])){
    total++;
    if(p.estado !== "cerrado"){ open++; if(p.fechaLimite && p.fechaLimite < todayStr()) overdue=true; }
  }
  return {open, total, overdue, eventos: h ? (h.eventos||[]).length : 0};
}

function histBtnHtml(val){
  const id = String(val);
  const ps = pendStats(id);
  const cls = ps.overdue ? "ovd" : (ps.open>0 ? "has" : "");
  return `<button class="hbtn ${cls}" data-eq="${esc(id)}" title="Historial / pendientes del equipo">📋<span class="hb">${ps.open>0?ps.open:""}</span></button>`;
}
function refreshBadges(){
  document.querySelectorAll(".hbtn").forEach(b=>{
    const ps = pendStats(b.dataset.eq);
    b.classList.toggle("has", ps.open>0 && !ps.overdue);
    b.classList.toggle("ovd", ps.overdue);
    const s = b.querySelector(".hb"); if(s) s.textContent = ps.open>0 ? ps.open : "";
  });
}

function getEquipoInfo(id){
  const r = REG_BY_ID[String(id)]; if(!r) return {};
  const g = k => (REG_COL[k]!=null ? r[REG_COL[k]] : "") || "";
  return { equipo:g("Equipo"), inv:g("N° Inventario"), marca:g("Marca"), modelo:g("Modelo"),
           serie:g("Serie"), servicio:g("Servicio"), ubic:g("Ubicación") };
}

function syncToolbar(){
  $("#globalSearch").value = state[CURRENT].search || "";
  const hf = $("#histFilter");
  hf.value = state[CURRENT].histFilter || "all";
  hf.style.display = (DATA[CURRENT].action === false) ? "none" : "";
}

/* ---- metadatos de vistas + vista agregada "Pendientes" ---- */
function computeMeta(){
  ["registro","pmp"].forEach(k=>{
    const v = DATA[k];
    v.action = true;
    v.unit = "equipos";
    v.idCol = v.columns.findIndex(c => !c.group && String(c.label).trim()==="ID");
  });
  const idc = DATA.registro.idCol;
  DATA.registro.columns.forEach((c,i)=>{ if(!c.group) REG_COL[String(c.label).trim()] = i; });
  DATA.registro.rows.forEach(r=>{ REG_BY_ID[String(r[idc])] = r; });

  DATA.pendientes = {
    id:"pendientes", title:"📋 Pendientes", action:false, idCol:-1, stickyCols:1, unit:"pendientes",
    columns:[
      {label:"Equipo",group:null,type:"text"},
      {label:"Fecha evento",group:null,type:"text"},
      {label:"Tipo evento",group:null,type:"text"},
      {label:"Pendiente",group:null,type:"text"},
      {label:"Estado",group:null,type:"status"},
      {label:"Responsable",group:null,type:"text"},
      {label:"Fecha límite",group:null,type:"text"},
      {label:"Tareas",group:null,type:"text"}
    ],
    rows:[], _eq:[]
  };
  state.pendientes = {filters:{}, sort:null, search:"", histFilter:"all"};
}

function rebuildPendientes(){
  const v = DATA.pendientes; v.rows = []; v._eq = [];
  Object.keys(HIST).forEach(id=>{
    const info = getEquipoInfo(id);
    (HIST[id].eventos||[]).forEach(ev=>{
      (ev.pendientes||[]).forEach(p=>{
        const ts = p.tareas||[];
        const hechas = ts.filter(t=>t.estado==="hecha").length;
        const ovd = p.estado!=="cerrado" && p.fechaLimite && p.fechaLimite < todayStr();
        v.rows.push([
          `#${id} ${info.equipo||""}`.trim(), ev.fecha||"", ev.tipo||"", p.titulo||"",
          p.estado||"", p.responsable||"", (p.fechaLimite||"") + (ovd?" ⚠":""), `${hechas}/${ts.length}`
        ]);
        v._eq.push(String(id));
      });
    });
  });
}

/* ---- panel lateral ---- */
function openDrawer(id){ CURRENT_EQ = String(id); editing = null; renderDrawer();
  $("#overlay").classList.add("show"); $("#drawer").classList.add("show"); }
function closeDrawer(){ $("#overlay").classList.remove("show"); $("#drawer").classList.remove("show");
  CURRENT_EQ = null; editing = null; }

function estadoChip(e){ const m={"abierto":"abierto","en proceso":"proceso","cerrado":"cerrado"};
  return `<span class="chip ${m[e]||"abierto"}">${esc(e||"abierto")}</span>`; }
function tipoChip(t){ return `<span class="chip ${t==="gestion"?"gestion":"tarea"}">${t==="gestion"?"Gestión":"Tarea"}</span>`; }
function optList(arr,sel){ return arr.map(o=>`<option ${o===sel?"selected":""}>${esc(o)}</option>`).join(""); }

function eventForm(ev){ const e=ev||{};
  return `<div class="form">
    <div class="frow">
      <div><label>Fecha</label><input type="date" id="f_fecha" value="${esc(e.fecha||todayStr())}"></div>
      <div><label>Tipo</label><select id="f_tipo">${optList(TIPOS_EVENTO, e.tipo||"Observación")}</select></div>
    </div>
    <label>¿Qué ocurrió?</label><textarea id="f_desc" rows="2" placeholder="Descripción del evento...">${esc(e.descripcion||"")}</textarea>
    <div class="facts"><button class="btn primary" data-act="save-event" data-ev="${e.id||""}">Guardar</button>
      <button class="btn" data-act="cancel">Cancelar</button></div></div>`;
}
function pendForm(evId,p){ const x=p||{};
  return `<div class="form">
    <label>Pendiente</label><input id="f_pt" value="${esc(x.titulo||"")}" placeholder="Ej: Reemplazar sensor">
    <div class="frow">
      <div><label>Responsable</label><input id="f_presp" value="${esc(x.responsable||"")}"></div>
      <div><label>Fecha límite</label><input type="date" id="f_pfl" value="${esc(x.fechaLimite||"")}"></div>
    </div>
    <label>Estado</label><select id="f_pest">${optList(["abierto","en proceso","cerrado"], x.estado||"abierto")}</select>
    <div class="facts"><button class="btn primary" data-act="save-pend" data-ev="${evId}" data-p="${x.id||""}">Guardar</button>
      <button class="btn" data-act="cancel">Cancelar</button></div></div>`;
}
function taskForm(evId,pId,t){ const x=t||{};
  return `<div class="form">
    <div class="frow">
      <div><label>Tipo</label><select id="f_ttipo"><option value="tarea" ${x.tipo!=="gestion"?"selected":""}>Tarea</option><option value="gestion" ${x.tipo==="gestion"?"selected":""}>Gestión</option></select></div>
      <div><label>Estado</label><select id="f_test"><option value="pendiente" ${x.estado!=="hecha"?"selected":""}>Pendiente</option><option value="hecha" ${x.estado==="hecha"?"selected":""}>Hecha</option></select></div>
    </div>
    <label>Descripción</label><input id="f_tdesc" value="${esc(x.descripcion||"")}" placeholder="Ej: Cotizar repuesto con proveedor">
    <label>Nota (opcional)</label><input id="f_tnota" value="${esc(x.nota||"")}">
    <div class="facts"><button class="btn primary" data-act="save-task" data-ev="${evId}" data-p="${pId}" data-t="${x.id||""}">Guardar</button>
      <button class="btn" data-act="cancel">Cancelar</button></div></div>`;
}

function renderDrawer(){
  const id = CURRENT_EQ; const rec = getHist(id); const info = getEquipoInfo(id);
  $("#dTitle").textContent = `Equipo #${id}${info.equipo?(" · "+info.equipo):""}`;
  $("#dSub").innerHTML = [
    info.inv?("N° Inv: "+esc(info.inv)):"", (info.marca||info.modelo)?esc((info.marca||"")+" "+(info.modelo||"")).trim():"",
    info.serie?("Serie: "+esc(info.serie)):"", info.servicio?esc(info.servicio):"", info.ubic?esc(info.ubic):""
  ].filter(Boolean).join(" · ");

  let b = "";
  b += (editing && editing.kind==="event" && editing.mode==="add")
        ? eventForm(null)
        : `<button class="addbtn" data-act="add-event">＋ Nuevo evento</button>`;

  const evs = (rec.eventos||[]).slice().reverse();   // más reciente primero
  if(!evs.length && !(editing&&editing.kind==="event"&&editing.mode==="add"))
    b += `<div class="dhint">Aún no hay eventos registrados para este equipo.</div>`;

  evs.forEach(ev=>{
    b += `<div class="ev">`;
    if(editing && editing.kind==="event" && editing.mode==="edit" && editing.evId===ev.id){
      b += eventForm(ev);
    }else{
      b += `<div class="evh"><span class="evdate">📅 ${esc(ev.fecha||"")} ${ev.tipo?`· <span class="chip tipoev">${esc(ev.tipo)}</span>`:""}</span>
        <span><button class="minibtn" data-act="edit-event" data-ev="${ev.id}" title="Editar">✏️</button>
        <button class="minibtn" data-act="del-event" data-ev="${ev.id}" title="Eliminar">🗑️</button></span></div>
        <div class="evdesc">${esc(ev.descripcion||"")}</div>`;
    }
    (ev.pendientes||[]).forEach(p=>{
      const ovd = p.estado!=="cerrado" && p.fechaLimite && p.fechaLimite < todayStr();
      b += `<div class="pend${ovd?" ovd":""}">`;
      if(editing && editing.kind==="pend" && editing.mode==="edit" && editing.pId===p.id){
        b += pendForm(ev.id, p);
      }else{
        b += `<div class="ph"><span class="pt">📌 ${esc(p.titulo||"")}</span>
          <span>${estadoChip(p.estado)}
          <button class="minibtn" data-act="edit-pend" data-ev="${ev.id}" data-p="${p.id}" title="Editar">✏️</button>
          <button class="minibtn" data-act="del-pend" data-ev="${ev.id}" data-p="${p.id}" title="Eliminar">🗑️</button></span></div>
          <div class="meta">${p.responsable?("👤 "+esc(p.responsable)+" "):""}${p.fechaLimite?("· ⏰ "+esc(p.fechaLimite)+(ovd?" (vencido)":"")):""}</div>`;
      }
      (p.tareas||[]).forEach(t=>{
        if(editing && editing.kind==="task" && editing.mode==="edit" && editing.tId===t.id){
          b += taskForm(ev.id, p.id, t);
        }else{
          b += `<div class="task${t.estado==="hecha"?" done":""}">
            <input type="checkbox" data-act="toggle-task" data-ev="${ev.id}" data-p="${p.id}" data-t="${t.id}" ${t.estado==="hecha"?"checked":""}>
            <div class="tdesc">${tipoChip(t.tipo)} ${esc(t.descripcion||"")}${t.nota?`<div class="meta">${esc(t.nota)}</div>`:""}</div>
            <button class="minibtn" data-act="edit-task" data-ev="${ev.id}" data-p="${p.id}" data-t="${t.id}" title="Editar">✏️</button>
            <button class="minibtn" data-act="del-task" data-ev="${ev.id}" data-p="${p.id}" data-t="${t.id}" title="Eliminar">🗑️</button></div>`;
        }
      });
      b += (editing && editing.kind==="task" && editing.mode==="add" && editing.pId===p.id)
            ? taskForm(ev.id, p.id, null)
            : `<button class="addbtn sub" data-act="add-task" data-ev="${ev.id}" data-p="${p.id}">＋ Tarea / gestión</button>`;
      b += `</div>`;
    });
    b += (editing && editing.kind==="pend" && editing.mode==="add" && editing.evId===ev.id)
          ? pendForm(ev.id, null)
          : `<button class="addbtn sub" data-act="add-pend" data-ev="${ev.id}">＋ Pendiente</button>`;
    b += `</div>`;
  });
  $("#dBody").innerHTML = b;
}

function val(id){ const el=document.getElementById(id); return el?el.value.trim():""; }
function saveEvent(evId){
  const rec = ensureEq(CURRENT_EQ);
  const data = {fecha:val("f_fecha"), tipo:val("f_tipo"), descripcion:val("f_desc")};
  if(!data.descripcion && !data.fecha){ editing=null; return renderDrawer(); }
  if(evId){ Object.assign(rec.eventos.find(x=>x.id===evId), data); }
  else { rec.eventos.push({id:uid(), ...data, pendientes:[]}); }
  editing=null; saveHist(); renderDrawer();
}
function savePend(evId,pId){
  const ev = ensureEq(CURRENT_EQ).eventos.find(x=>x.id===evId); if(!ev) return;
  const data = {titulo:val("f_pt"), responsable:val("f_presp"), fechaLimite:val("f_pfl"), estado:val("f_pest")};
  if(!data.titulo){ editing=null; return renderDrawer(); }
  if(pId){ Object.assign(ev.pendientes.find(x=>x.id===pId), data); }
  else { (ev.pendientes=ev.pendientes||[]).push({id:uid(), ...data, tareas:[]}); }
  editing=null; saveHist(); renderDrawer();
}
function saveTask(evId,pId,tId){
  const ev = ensureEq(CURRENT_EQ).eventos.find(x=>x.id===evId); if(!ev) return;
  const p = ev.pendientes.find(x=>x.id===pId); if(!p) return;
  const data = {tipo:val("f_ttipo"), estado:val("f_test"), descripcion:val("f_tdesc"), nota:val("f_tnota")};
  if(!data.descripcion){ editing=null; return renderDrawer(); }
  if(tId){ Object.assign(p.tareas.find(x=>x.id===tId), data); }
  else { (p.tareas=p.tareas||[]).push({id:uid(), ...data}); }
  editing=null; saveHist(); renderDrawer();
}

const dBody = $("#dBody");
dBody.addEventListener("click", e=>{
  const b = e.target.closest("[data-act]"); if(!b) return;
  const {act, ev:evId, p:pId, t:tId} = b.dataset;
  if(act==="cancel"){ editing=null; return renderDrawer(); }
  if(act==="add-event"){ editing={kind:"event",mode:"add"}; return renderDrawer(); }
  if(act==="edit-event"){ editing={kind:"event",mode:"edit",evId}; return renderDrawer(); }
  if(act==="save-event"){ return saveEvent(evId); }
  if(act==="del-event"){ if(confirm("¿Eliminar este evento y todo su contenido?")){
      const rec=ensureEq(CURRENT_EQ); rec.eventos=rec.eventos.filter(x=>x.id!==evId); editing=null; saveHist(); renderDrawer(); } return; }
  if(act==="add-pend"){ editing={kind:"pend",mode:"add",evId}; return renderDrawer(); }
  if(act==="edit-pend"){ editing={kind:"pend",mode:"edit",evId,pId}; return renderDrawer(); }
  if(act==="save-pend"){ return savePend(evId,pId); }
  if(act==="del-pend"){ if(confirm("¿Eliminar este pendiente y sus tareas?")){
      const ev=ensureEq(CURRENT_EQ).eventos.find(x=>x.id===evId); ev.pendientes=ev.pendientes.filter(x=>x.id!==pId); editing=null; saveHist(); renderDrawer(); } return; }
  if(act==="add-task"){ editing={kind:"task",mode:"add",evId,pId}; return renderDrawer(); }
  if(act==="edit-task"){ editing={kind:"task",mode:"edit",evId,pId,tId}; return renderDrawer(); }
  if(act==="save-task"){ return saveTask(evId,pId,tId); }
  if(act==="del-task"){ const ev=ensureEq(CURRENT_EQ).eventos.find(x=>x.id===evId);
      const p=ev.pendientes.find(x=>x.id===pId); p.tareas=p.tareas.filter(x=>x.id!==tId); editing=null; saveHist(); renderDrawer(); return; }
});
dBody.addEventListener("change", e=>{
  const c = e.target.closest('[data-act="toggle-task"]'); if(!c) return;
  const ev = ensureEq(CURRENT_EQ).eventos.find(x=>x.id===c.dataset.ev);
  const p = ev.pendientes.find(x=>x.id===c.dataset.p);
  const t = p.tareas.find(x=>x.id===c.dataset.t);
  t.estado = c.checked ? "hecha" : "pendiente"; saveHist(); renderDrawer();
});
$("#dClose").addEventListener("click", closeDrawer);
$("#overlay").addEventListener("click", closeDrawer);
document.addEventListener("keydown", e=>{ if(e.key==="Escape" && $("#drawer").classList.contains("show")) closeDrawer(); });

/* ---- respaldo del historial (export / import JSON) ---- */
$("#expHist").addEventListener("click", ()=>{
  const blob = new Blob([JSON.stringify(HIST,null,2)], {type:"application/json"});
  const a = document.createElement("a"); a.href = URL.createObjectURL(blob);
  a.download = "historial_MP_2026_" + todayStr() + ".json"; a.click();
});
$("#impHist").addEventListener("click", ()=> $("#impFile").click());
$("#impFile").addEventListener("change", e=>{
  const f = e.target.files[0]; if(!f) return;
  const r = new FileReader();
  r.onload = ()=>{
    try{
      const obj = JSON.parse(r.result);
      if(typeof obj!=="object" || Array.isArray(obj)) throw 0;
      if(confirm("¿Reemplazar el historial actual por el del archivo? (Respalda antes con «⬇ Historial» si lo necesitas.)")){
        HIST = obj; saveHist();
        if(CURRENT_EQ) renderDrawer();
        renderTable();
      }
    }catch(err){ alert("El archivo no es un respaldo JSON válido."); }
    finally{ e.target.value = ""; }
  };
  r.readAsText(f);
});

/* ---------- init ---------- */
loadHist();
computeMeta();
buildTabs(); buildLegend(); renderTable();
</script>
</body>
</html>
"""

html_out = HTML.replace("__DATA__", data_json)
with open(OUT, "w", encoding="utf-8") as f:
    f.write(html_out)
print("HTML escrito en:", OUT, "tamaño:", len(html_out), "bytes")
