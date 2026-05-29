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
  <select id="histFilter" title="Filtrar por pendientes o estado del equipo">
    <option value="all">📋 Pendientes: todos</option>
    <option value="open">Con pendientes abiertos</option>
    <option value="overdue">Con pendientes vencidos</option>
    <optgroup label="Estado del equipo">
      <option value="est:operativo">Operativo</option>
      <option value="est:no operativo">No operativo</option>
      <option value="est:en servicio técnico">En servicio técnico</option>
      <option value="est:baja">Baja</option>
    </optgroup>
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
  ["sw-baja","Baja"],["sw-no","No realizada"],["sw-nu","NU – No utilizado"],
  ["est-operativo","● Equipo operativo"],["est-nooperativo","● No operativo"],
  ["est-servicio","● En servicio técnico"],["est-baja","● Baja"]
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
    const eqId = isEq ? String(row[v.idCol]) : "";
    const red = isEq ? reduceEquipo(eqId) : null;   // estado + overlay R (mes sig.)
    let tds = "";
    cols.forEach((c,idx) => {
      let value = row[idx] ?? "";
      let cls = "", inner, ovl = false;
      if(c.type === "status"){
        // overlay "R" generado por evento (C1–C8) si la celda del mes está vacía
        if(red && v.overlayCol && v.overlayCol[idx]!=null && value==="" && red.progR.has(v.overlayCol[idx])){ value="R"; ovl=true; }
        const sc = statusClass(value); cls = "st" + (sc?(" "+sc):"") + (isEq?" mes":"") + (ovl?" ovl":"");
        inner = esc(value);
      }else{
        inner = `<div class='cell' title="${esc(value)}">${esc(value)}</div>`;
      }
      // botón de historial + estado en la columna ID (solo vistas de equipos)
      if(isEq && idx === v.idCol){ inner = histBtnHtml(value, red) + inner; }
      let style = "";
      if(idx < sticky){ cls += (cls?" ":"") + "sticky"; style = leftStyle(cols,idx,sticky); }
      const ttl = ovl ? ` title="Reprogramado por evento de MP"` : "";
      tds += `<td class='${cls}' style='${style}'${ttl}>${inner}</td>`;
    });
    // en la vista Pendientes la fila completa abre el historial del equipo
    const clk = (v.action===false) ? " class='clickrow'" : "";
    return `<tr data-ri='${ri}'${clk}>${tds}</tr>`;
  }).join("");

  const _sl = boxEl.scrollLeft, _st = boxEl.scrollTop;     // conservar scroll
  boxEl.innerHTML = `<table>${thead}<tbody>${rowsHtml}</tbody></table>`;
  boxEl.scrollLeft = _sl; boxEl.scrollTop = _st;
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
    // filtro por estado de pendientes o estado del equipo (solo vistas de equipos)
    if(ok && isEq && hf !== "all"){
      const id = String(row[v.idCol]);
      if(hf === "open"){ if(pendStats(id).open <= 0) ok = false; }
      else if(hf === "overdue"){ if(!pendStats(id).overdue) ok = false; }
      else if(hf.startsWith("est:")){ if(equipoEstado(id) !== hf.slice(4)) ok = false; }
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
   MOTOR DE EVENTOS POR EQUIPO
   7 tipos de evento -> estado del equipo + ciclos correctivos (Folio SIGEM).
   El estado y los ciclos se DERIVAN del registro de eventos (la anulación
   recalcula todo). Pendientes a nivel de equipo (auto + manuales) con tareas.
   Persistencia: localStorage + respaldo JSON.
   =================================================================== */
const HKEY = "mp2026_historial_v2";
let HIST = {};
let CURRENT_EQ = null;        // equipo abierto en el panel
let editing = null;           // formulario abierto
let REG_COL = {}, REG_BY_ID = {};
const MESES_AB = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"];

const ETIPOS = ["Solicitud de trabajo","Visita técnica","Orden de Compra",
  "Envío a servicio técnico","Recepción","Reparación","Mantención preventiva"];
const MP_RESULT = ["Si","C1","C2","C3","C4","C5","C6","C7","C8","FS","Baja","NU","No"];
const MP_RESULT_TXT = {
  "Si":"MP realizada","C1":"Imposibilidad de desocupar el equipo","C2":"En servicio técnico",
  "C3":"No operativo, espera de repuestos","C4":"En préstamo a otro hospital",
  "C5":"No disponibilidad HH SEC","C6":"No disponibilidad ST externo",
  "C7":"Ausencia funcionario SEC > 15 días","C8":"Contingencia hospitalaria",
  "FS":"Fuera de servicio","Baja":"Equipo dado de baja","NU":"No ubicable","No":"No realizada"
};
const ESTADOS = ["operativo","no operativo","en servicio técnico","baja"];
function mpEstado(code){
  if(code==="Si") return "operativo";
  if(code==="C2") return "en servicio técnico";
  if(code==="C3"||code==="FS"||code==="NU") return "no operativo";
  if(code==="Baja") return "baja";
  return null;   // C1, C4-C8, No -> sin cambio
}
function mapEstadoLabel(l){ return {"Operativo":"operativo","No operativo":"no operativo",
  "En servicio técnico":"en servicio técnico"}[l] || l; }
function estCls(e){ return {"operativo":"est-operativo","no operativo":"est-nooperativo",
  "en servicio técnico":"est-servicio","baja":"est-baja"}[e] || "est-operativo"; }

function todayStr(){ const d=new Date(); return new Date(d.getTime()-d.getTimezoneOffset()*60000).toISOString().slice(0,10); }
function uid(){ return "x"+Math.random().toString(36).slice(2,9)+Date.now().toString(36).slice(-4); }
function monthIdx(f){ const m=/^\d{4}-(\d{2})/.exec(f||""); return m?(+m[1]-1):new Date().getMonth(); }
function eqIds(){ return Object.keys(HIST).filter(k=>!k.startsWith("__")); }

function loadHist(){
  try{ HIST = JSON.parse(localStorage.getItem(HKEY) || "{}") || {}; }catch(e){ HIST = {}; }
  // migración v1 (pendientes anidados bajo eventos) -> v2 (pendientes por equipo)
  try{
    const old = JSON.parse(localStorage.getItem("mp2026_historial_v1") || "null");
    if(old && !eqIds().length){
      for(const id in old){ const e=old[id]||{}; HIST[id]={eventos:e.eventos||[],pendientes:e.pendientes||[]};
        (HIST[id].eventos||[]).forEach(ev=>{ if(ev.pendientes){ ev.pendientes.forEach(p=>{p.originEvId=ev.id;HIST[id].pendientes.push(p);}); delete ev.pendientes; } }); }
    }
  }catch(e){}
  eqIds().forEach(id=>{ HIST[id].eventos=HIST[id].eventos||[]; HIST[id].pendientes=HIST[id].pendientes||[]; });
}
function saveHist(){
  try{ localStorage.setItem(HKEY, JSON.stringify(HIST)); }
  catch(e){ alert("No se pudo guardar (almacenamiento lleno o bloqueado). Usa «⬇ Historial» para respaldar."); }
  afterChange();
}
function afterChange(){
  refreshBadges();
  renderTable();              // refresca overlays/estado/pendientes (conserva scroll)
  if(CURRENT_EQ) renderDrawer();
}
function getHist(id){ const h=HIST[id]||{eventos:[],pendientes:[]}; h.eventos=h.eventos||[]; h.pendientes=h.pendientes||[]; return h; }
function ensureEq(id){ if(!HIST[id]) HIST[id]={eventos:[],pendientes:[]};
  HIST[id].eventos=HIST[id].eventos||[]; HIST[id].pendientes=HIST[id].pendientes||[]; return HIST[id]; }

/* ---- REDUCER: estado actual, ciclos SIGEM abiertos y overlay R (mes sig.) ---- */
function reduceEquipo(id){
  const h = HIST[id]; let estado="operativo"; const open=new Map(); const progR=new Set();
  if(h && h.eventos && h.eventos.length){
    const evs = h.eventos.filter(e=>!e.anulado).slice().sort((a,b)=>
      (a.fecha||"").localeCompare(b.fecha||"") || (a.seq||0)-(b.seq||0));
    for(const e of evs){
      switch(e.tipo){
        case "Solicitud de trabajo":
          estado="no operativo"; if(e.folio) open.set(e.folio,{folio:e.folio,fecha:e.fecha}); break;
        case "Visita técnica":
          estado=mapEstadoLabel(e.estadoResultante)||estado;
          if(e.tipoVisita==="Correctiva" && mapEstadoLabel(e.estadoResultante)==="operativo" && e.folio) open.delete(e.folio);
          break;
        case "Orden de Compra": break;
        case "Envío a servicio técnico": estado="en servicio técnico"; break;
        case "Recepción":
          estado=mapEstadoLabel(e.estadoResultante)||estado;
          if(mapEstadoLabel(e.estadoResultante)==="operativo" && e.folio) open.delete(e.folio); break;
        case "Reparación":
          estado=mapEstadoLabel(e.estadoResultante)||estado;
          if(mapEstadoLabel(e.estadoResultante)==="operativo" && e.folio) open.delete(e.folio); break;
        case "Mantención preventiva": {
          const ne=mpEstado(e.resultado); if(ne) estado=ne;
          if(/^C[1-8]$/.test(e.resultado||"")) progR.add((monthIdx(e.fecha)+1)%12);
          break;
        }
      }
    }
  }
  return {estado, open, progR};
}
function equipoEstado(id){ return reduceEquipo(String(id)).estado; }
function openCycles(id){ return [...reduceEquipo(String(id)).open.values()]; }

function pendStats(id){
  const h=HIST[id]; let open=0,total=0,overdue=false;
  if(h) for(const p of (h.pendientes||[])){ if(p.anulado) continue; total++;
    if(p.estado!=="cerrado"){ open++; if(p.fechaLimite && p.fechaLimite<todayStr()) overdue=true; } }
  return {open,total,overdue};
}

function histBtnHtml(value, red){
  const id=String(value); const ps=pendStats(id);
  const est=(red?red.estado:equipoEstado(id));
  const cls = ps.overdue ? "ovd" : (ps.open>0 ? "has" : "");
  const dot = est!=="operativo" ? `<span class="estdot ${estCls(est)}" title="${esc(est)}"></span>` : "";
  return `<button class="hbtn ${cls}" data-eq="${esc(id)}" title="Historial · estado: ${esc(est)}">📋<span class="hb">${ps.open>0?ps.open:""}</span></button>${dot}`;
}
function refreshBadges(){
  document.querySelectorAll(".hbtn").forEach(b=>{
    const ps=pendStats(b.dataset.eq);
    b.classList.toggle("has", ps.open>0 && !ps.overdue); b.classList.toggle("ovd", ps.overdue);
    const s=b.querySelector(".hb"); if(s) s.textContent=ps.open>0?ps.open:"";
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
    const v = DATA[k]; v.action=true; v.unit="equipos";
    v.idCol = v.columns.findIndex(c => !c.group && String(c.label).trim()==="ID");
    // mapa colIndex -> mesIndex para overlay "R" (Registro: 2ª de cada par P/R; PMP: cada mes)
    v.overlayCol = {}; let s=0;
    v.columns.forEach((c,i)=>{ if(c.type==="status"){
      if(k==="registro"){ if(s%2===1) v.overlayCol[i]=Math.floor(s/2); }
      else { v.overlayCol[i]=s; }
      s++;
    }});
  });
  const idc = DATA.registro.idCol;
  DATA.registro.columns.forEach((c,i)=>{ if(!c.group) REG_COL[String(c.label).trim()] = i; });
  DATA.registro.rows.forEach(r=>{ REG_BY_ID[String(r[idc])] = r; });

  DATA.pendientes = {
    id:"pendientes", title:"📋 Pendientes", action:false, idCol:-1, stickyCols:1, unit:"pendientes",
    columns:[
      {label:"Equipo",group:null,type:"text"},
      {label:"Pendiente",group:null,type:"text"},
      {label:"Tipo",group:null,type:"text"},
      {label:"Estado",group:null,type:"status"},
      {label:"Responsable",group:null,type:"text"},
      {label:"Fecha límite",group:null,type:"text"},
      {label:"Tareas",group:null,type:"text"},
      {label:"Folio SIGEM",group:null,type:"text"}
    ],
    rows:[], _eq:[]
  };
  state.pendientes = {filters:{}, sort:null, search:"", histFilter:"all"};
}

function rebuildPendientes(){
  const v = DATA.pendientes; v.rows = []; v._eq = [];
  eqIds().forEach(id=>{
    const info = getEquipoInfo(id);
    (HIST[id].pendientes||[]).forEach(p=>{
      if(p.anulado) return;
      const ts = p.tareas||[]; const hechas = ts.filter(t=>t.estado==="hecha").length;
      const ovd = p.estado!=="cerrado" && p.fechaLimite && p.fechaLimite < todayStr();
      v.rows.push([
        `#${id} ${info.equipo||""}`.trim(), p.titulo||"", p.tipo||"", p.estado||"",
        p.responsable||"", (p.fechaLimite||"") + (ovd?" ⚠":""), `${hechas}/${ts.length}`, p.folio||""
      ]);
      v._eq.push(String(id));
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
function fLab(t){ return `<label>${esc(t)}</label>`; }
function fIn(id,v,ph,type){ return `<input id="${id}" type="${type||"text"}" value="${esc(v||"")}" placeholder="${esc(ph||"")}">`; }
function fSel(id,opts,sel){ return `<select id="${id}">${opts.map(o=>`<option ${o===sel?"selected":""}>${esc(o)}</option>`).join("")}</select>`; }
function folioSel(id,cycles,val){
  const opts=cycles.map(c=>c.folio); if(val && !opts.includes(val)) opts.unshift(val);
  if(!opts.length) return `<div class="ro">Sin ciclo SIGEM abierto.</div><input id="${id}" type="hidden" value="">`;
  const cur = val||opts[0];
  return `<select id="${id}">${opts.map(o=>`<option ${o===cur?"selected":""}>${esc(o)}</option>`).join("")}</select>`;
}

function autoFolio(){
  let max=0; eqIds().forEach(id=>(HIST[id].eventos||[]).forEach(e=>{
    const m=/SIGEM-2026-(\d+)/.exec(e.folio||""); if(m) max=Math.max(max,+m[1]); }));
  return "SIGEM-2026-"+String(max+1).padStart(4,"0");
}
function nextEnvio(){
  let max=0; eqIds().forEach(id=>(HIST[id].eventos||[]).forEach(e=>{
    if(e.tipo==="Envío a servicio técnico"){ const n=parseInt(e.nEnvio,10); if(!isNaN(n)) max=Math.max(max,n); } }));
  return max+1;
}

function eventFields(tipo,e,ctx){
  const ta=(id,v)=>`<textarea id="${id}" rows="2">${esc(v||"")}</textarea>`;
  const cyc=()=>folioSel("f_folio",ctx.cycles,e.folio);
  switch(tipo){
    case "Solicitud de trabajo":
      return fLab("Folio SIGEM (vacío = automático)")+fIn("f_folio",e.folio,"SIGEM-2026-####")
           + fLab("Descripción de la falla")+ta("f_obs",e.obs);
    case "Visita técnica":
      return `<div class="frow"><div>${fLab("Empresa")+fIn("f_empresa",e.empresa)}</div><div>${fLab("Técnico")+fIn("f_tecnico",e.tecnico)}</div></div>`
           + fLab("Tipo de visita")+fSel("f_tvisita",["Diagnóstica","Correctiva"],e.tipoVisita||"Diagnóstica")
           + fLab("Vincular a Folio SIGEM")+cyc()
           + fLab("Estado resultante")+fSel("f_estado",["No operativo","Operativo","En servicio técnico"],e.estadoResultante||"No operativo")
           + fLab("Informe / Observación")+ta("f_obs",e.obs);
    case "Orden de Compra":
      return `<div class="frow"><div>${fLab("N° Cotización")+fIn("f_cot",e.nCotizacion)}</div><div>${fLab("N° OC")+fIn("f_oc",e.nOC)}</div></div>`
           + fLab("Empresa")+fIn("f_empresa",e.empresa)
           + fLab("Vía")+fSel("f_via",["Trato directo","Compra ágil"],e.via||"Trato directo")
           + `<div id="wrap_finforme" style="display:${(e.via||"Trato directo")==="Trato directo"?"":"none"}">`+fLab("Folio informe (Trato directo)")+fIn("f_finforme",e.folioInforme)+`</div>`
           + fLab("Vincular a Folio SIGEM")+cyc()
           + fLab("Observación")+ta("f_obs",e.obs);
    case "Envío a servicio técnico":
      return fLab("N° de envío (vacío = correlativo)")+fIn("f_nenvio",e.nEnvio)
           + fLab("Empresa")+fIn("f_empresa",e.empresa)
           + fLab("Vincular a Folio SIGEM")+cyc()
           + `<div class="ro">Estado resultante: <b>en servicio técnico</b> (fijo)</div>`
           + fLab("Observación")+ta("f_obs",e.obs);
    case "Recepción":
      return fLab("N° envío original")+fIn("f_nenvio",e.nEnvioOriginal)
           + fLab("Folio guía de despacho")+fIn("f_guia",e.folioGuia)
           + fLab("Vincular a Folio SIGEM")+cyc()
           + fLab("Estado resultante")+fSel("f_estado",["No operativo","Operativo"],e.estadoResultante||"Operativo")
           + fLab("Informe técnico / Observación")+ta("f_obs",e.obs);
    case "Reparación":
      return fLab("Repuestos")+ta("f_repuestos",e.repuestos)
           + fLab("Vincular a Folio SIGEM")+cyc()
           + fLab("Estado resultante")+fSel("f_estado",["Operativo","No operativo","En servicio técnico"],e.estadoResultante||"Operativo")
           + fLab("Descripción de la tarea")+ta("f_obs",e.obs);
    case "Mantención preventiva": {
      const est0=mpEstado(e.resultado||"Si");
      return fLab("Ejecutor 2 (opcional)")+fIn("f_ejecutor2",e.ejecutor2)
           + fLab("Resultado")+fSel("f_resultado",MP_RESULT,e.resultado||"Si")
           + `<div class="ro">Estado resultante: <b id="mpEst">${est0?esc(est0):"(sin cambio)"}</b> (automático)</div>`
           + fLab("Observación")+ta("f_obs",e.obs);
    }
  }
  return "";
}

function eventForm(ev){
  const isEdit=!!ev; const tipo = isEdit ? ev.tipo : (editing.tipo||ETIPOS[0]);
  const cycles = openCycles(CURRENT_EQ);
  let h = `<div class="form">`;
  h += isEdit ? `<div class="evkv"><b>Tipo:</b> <span class="evtipo">${esc(tipo)}</span></div>`
              : fLab("Tipo de evento")+fSel("f_tipo",ETIPOS,tipo);
  h += fLab("Fecha")+fIn("f_fecha",(ev&&ev.fecha)||todayStr(),"","date");
  h += eventFields(tipo, ev||{}, {cycles});
  h += `<div class="facts"><button class="btn primary" data-act="save-event" data-ev="${isEdit?ev.id:""}">Guardar</button>`
     + `<button class="btn" data-act="cancel">Cancelar</button></div></div>`;
  return h;
}
function pendForm(p){ const x=p||{};
  return `<div class="form">
    ${fLab("Pendiente")+fIn("f_pt",x.titulo,"Ej: Localizar equipo")}
    ${fLab("Tipo")+fIn("f_ptipo",x.tipo||"Manual")}
    <div class="frow"><div>${fLab("Responsable")+fIn("f_presp",x.responsable)}</div><div>${fLab("Fecha límite")+fIn("f_pfl",x.fechaLimite,"","date")}</div></div>
    ${fLab("Estado")+fSel("f_pest",["abierto","en proceso","cerrado"],x.estado||"abierto")}
    <div class="facts"><button class="btn primary" data-act="save-pend" data-p="${x.id||""}">Guardar</button>
      <button class="btn" data-act="cancel">Cancelar</button></div></div>`;
}
function taskForm(pId,t){ const x=t||{};
  return `<div class="form">
    <div class="frow"><div>${fLab("Tipo")}<select id="f_ttipo"><option value="tarea" ${x.tipo!=="gestion"?"selected":""}>Tarea</option><option value="gestion" ${x.tipo==="gestion"?"selected":""}>Gestión</option></select></div>
    <div>${fLab("Estado")}<select id="f_test"><option value="pendiente" ${x.estado!=="hecha"?"selected":""}>Pendiente</option><option value="hecha" ${x.estado==="hecha"?"selected":""}>Hecha</option></select></div></div>
    ${fLab("Descripción")+fIn("f_tdesc",x.descripcion,"Ej: Cotizar repuesto")}
    ${fLab("Nota (opcional)")+fIn("f_tnota",x.nota)}
    <div class="facts"><button class="btn primary" data-act="save-task" data-p="${pId}" data-t="${x.id||""}">Guardar</button>
      <button class="btn" data-act="cancel">Cancelar</button></div></div>`;
}

function evSummary(e){
  const kv=[]; const push=(k,v)=>{ if(v) kv.push(`<b>${esc(k)}:</b> ${esc(v)}`); };
  switch(e.tipo){
    case "Solicitud de trabajo": push("Falla",e.obs); break;
    case "Visita técnica": push("Empresa",e.empresa); push("Técnico",e.tecnico); push("Visita",e.tipoVisita); push("Resultado",e.estadoResultante); push("Informe",e.obs); break;
    case "Orden de Compra": push("N° Cotización",e.nCotizacion); push("N° OC",e.nOC); push("Empresa",e.empresa); push("Vía",e.via); push("Folio informe",e.folioInforme); push("Obs",e.obs); break;
    case "Envío a servicio técnico": push("N° envío",e.nEnvio); push("Empresa",e.empresa); push("Obs",e.obs); break;
    case "Recepción": push("N° envío orig.",e.nEnvioOriginal); push("Guía despacho",e.folioGuia); push("Resultado",e.estadoResultante); push("Informe",e.obs); break;
    case "Reparación": push("Repuestos",e.repuestos); push("Resultado",e.estadoResultante); push("Tarea",e.obs); break;
    case "Mantención preventiva": push("Resultado",`${e.resultado}${MP_RESULT_TXT[e.resultado]?(" · "+MP_RESULT_TXT[e.resultado]):""}`); push("Ejecutor 2",e.ejecutor2); push("Obs",e.obs); break;
  }
  return kv.join("<br>") || "<i>—</i>";
}

function renderDrawer(){
  const id = CURRENT_EQ; const rec = getHist(id); const info = getEquipoInfo(id);
  const red = reduceEquipo(id);
  $("#dTitle").innerHTML = `Equipo #${esc(id)}${info.equipo?(" · "+esc(info.equipo)):""} <span class="est ${estCls(red.estado)}">${esc(red.estado)}</span>`;
  $("#dSub").innerHTML = [
    info.inv?("N° Inv: "+esc(info.inv)):"", (info.marca||info.modelo)?esc((info.marca||"")+" "+(info.modelo||"")).trim():"",
    info.serie?("Serie: "+esc(info.serie)):"", info.servicio?esc(info.servicio):"", info.ubic?esc(info.ubic):""
  ].filter(Boolean).join(" · ");

  const cyc=[...red.open.values()];
  let b = `<div class="evkv">Ciclos SIGEM abiertos: ${cyc.length?cyc.map(c=>`<span class="folio">${esc(c.folio)}</span>`).join(" "):"<i>ninguno</i>"}</div>`;

  // ---- EVENTOS ----
  b += `<div class="sec">Eventos</div>`;
  b += (editing && editing.kind==="event" && editing.mode==="add")
        ? eventForm(null) : `<button class="addbtn" data-act="add-event">＋ Nuevo evento</button>`;
  const evs = (rec.eventos||[]).slice().sort((a,b)=>(b.fecha||"").localeCompare(a.fecha||"")||(b.seq||0)-(a.seq||0));
  if(!evs.length && !(editing&&editing.kind==="event"&&editing.mode==="add")) b += `<div class="dhint">Sin eventos registrados.</div>`;
  evs.forEach(ev=>{
    if(editing && editing.kind==="event" && editing.mode==="edit" && editing.evId===ev.id){
      b += `<div class="ev">${eventForm(ev)}</div>`; return;
    }
    b += `<div class="ev${ev.anulado?" anul":""}">
      <div class="evh"><span class="evdate">📅 ${esc(ev.fecha||"")} <span class="evtipo">${esc(ev.tipo)}</span> ${ev.folio?`<span class="folio">${esc(ev.folio)}</span>`:""}</span>
      <span>${ev.anulado?"":`<button class="minibtn" data-act="edit-event" data-ev="${ev.id}" title="Editar">✏️</button><button class="minibtn" data-act="anul-event" data-ev="${ev.id}" title="Anular">🚫</button>`}</span></div>
      <div class="evkv">${evSummary(ev)}</div>
      ${ev.anulado?`<div class="anulmsg">Anulado (${esc(ev.anulFecha||"")}): ${esc(ev.anulMotivo||"")}</div>`:""}</div>`;
  });

  // ---- PENDIENTES ----
  b += `<div class="sec">Pendientes</div>`;
  b += (editing && editing.kind==="pend" && editing.mode==="add")
        ? pendForm(null) : `<button class="addbtn" data-act="add-pend">＋ Nuevo pendiente</button>`;
  const pends = (rec.pendientes||[]).filter(p=>!p.anulado);
  if(!pends.length && !(editing&&editing.kind==="pend"&&editing.mode==="add")) b += `<div class="dhint">Sin pendientes.</div>`;
  pends.forEach(p=>{
    const ovd = p.estado!=="cerrado" && p.fechaLimite && p.fechaLimite < todayStr();
    b += `<div class="pend${ovd?" ovd":""}">`;
    if(editing && editing.kind==="pend" && editing.mode==="edit" && editing.pId===p.id){ b += pendForm(p); }
    else{
      b += `<div class="ph"><span class="pt">📌 ${esc(p.titulo||"")} ${p.auto?`<span class="pauto">auto</span>`:""}</span>
        <span>${estadoChip(p.estado)} <button class="minibtn" data-act="edit-pend" data-p="${p.id}" title="Editar">✏️</button>
        <button class="minibtn" data-act="del-pend" data-p="${p.id}" title="Eliminar">🗑️</button></span></div>
        <div class="meta">${p.tipo?("· "+esc(p.tipo)+" "):""}${p.responsable?("· 👤 "+esc(p.responsable)+" "):""}${p.fechaLimite?("· ⏰ "+esc(p.fechaLimite)+(ovd?" (vencido)":"")):""}${p.folio?(' · <span class="folio">'+esc(p.folio)+"</span>"):""}</div>`;
    }
    (p.tareas||[]).forEach(t=>{
      if(editing && editing.kind==="task" && editing.mode==="edit" && editing.tId===t.id){ b += taskForm(p.id,t); }
      else b += `<div class="task${t.estado==="hecha"?" done":""}">
        <input type="checkbox" data-act="toggle-task" data-p="${p.id}" data-t="${t.id}" ${t.estado==="hecha"?"checked":""}>
        <div class="tdesc">${tipoChip(t.tipo)} ${esc(t.descripcion||"")}${t.nota?`<div class="meta">${esc(t.nota)}</div>`:""}</div>
        <button class="minibtn" data-act="edit-task" data-p="${p.id}" data-t="${t.id}" title="Editar">✏️</button>
        <button class="minibtn" data-act="del-task" data-p="${p.id}" data-t="${t.id}" title="Eliminar">🗑️</button></div>`;
    });
    b += (editing && editing.kind==="task" && editing.mode==="add" && editing.pId===p.id)
          ? taskForm(p.id,null) : `<button class="addbtn sub" data-act="add-task" data-p="${p.id}">＋ Tarea / gestión</button>`;
    b += `</div>`;
  });
  $("#dBody").innerHTML = b;
}

function val(id){ const el=document.getElementById(id); return el?el.value.trim():""; }

function saveEvent(evId){
  const rec = ensureEq(CURRENT_EQ);
  const existing = evId ? rec.eventos.find(x=>x.id===evId) : null;
  const tipo = existing ? existing.tipo : editing.tipo;
  const fecha = val("f_fecha");
  if(!fecha){ alert("Indica la fecha del evento."); return; }
  const cycles = openCycles(CURRENT_EQ);
  const data = {fecha, tipo, obs:val("f_obs")};
  if(tipo==="Solicitud de trabajo"){ data.folio = val("f_folio") || autoFolio(); }
  else if(tipo==="Visita técnica"){ data.empresa=val("f_empresa"); data.tecnico=val("f_tecnico"); data.tipoVisita=val("f_tvisita"); data.folio=val("f_folio"); data.estadoResultante=val("f_estado"); }
  else if(tipo==="Orden de Compra"){ data.nCotizacion=val("f_cot"); data.nOC=val("f_oc"); data.empresa=val("f_empresa"); data.via=val("f_via"); data.folioInforme=(data.via==="Trato directo")?val("f_finforme"):""; data.folio=val("f_folio"); }
  else if(tipo==="Envío a servicio técnico"){ data.nEnvio=val("f_nenvio")||String(nextEnvio()); data.empresa=val("f_empresa"); data.folio=val("f_folio"); }
  else if(tipo==="Recepción"){ data.nEnvioOriginal=val("f_nenvio"); data.folioGuia=val("f_guia"); data.folio=val("f_folio"); data.estadoResultante=val("f_estado");
      if(!cycles.length && !confirm("Este equipo no tiene un ciclo SIGEM abierto. ¿Registrar la recepción de todos modos?")) return; }
  else if(tipo==="Reparación"){ data.repuestos=val("f_repuestos"); data.folio=val("f_folio"); data.estadoResultante=val("f_estado");
      if(!cycles.length && !confirm("Este equipo no tiene un ciclo SIGEM abierto. ¿Registrar la reparación de todos modos?")) return; }
  else if(tipo==="Mantención preventiva"){ data.ejecutor2=val("f_ejecutor2"); data.resultado=val("f_resultado"); }

  let theId;
  if(existing){ theId=existing.id; Object.assign(existing, data); }
  else { theId=uid(); rec.eventos.push({id:theId, seq:Date.now(), origen:"manual", anulado:false, ...data}); }

  // pendientes automáticos de la MP (se regeneran)
  rec.pendientes = rec.pendientes.filter(p=>!(p.auto && p.originEvId===theId));
  if(tipo==="Mantención preventiva"){
    const code=data.resultado;
    if(/^C[1-8]$/.test(code)){
      const mes=MESES_AB[(monthIdx(fecha)+1)%12];
      rec.pendientes.push({id:uid(),auto:true,originEvId:theId,tipo:"Reprogramación",titulo:`Reprogramar MP (${code}) → ${mes}`,estado:"abierto",responsable:"",fechaLimite:"",folio:"",tareas:[]});
    }else if(code==="NU"){
      rec.pendientes.push({id:uid(),auto:true,originEvId:theId,tipo:"Localización",titulo:"Localizar equipo",estado:"abierto",responsable:"",fechaLimite:"",folio:"",tareas:[]});
    }
  }
  editing=null; saveHist();
}
function anularEvento(evId){
  const motivo = prompt("Motivo de la anulación:"); if(motivo===null) return;
  const rec = ensureEq(CURRENT_EQ); const ev = rec.eventos.find(x=>x.id===evId); if(!ev) return;
  ev.anulado=true; ev.anulFecha=todayStr(); ev.anulMotivo=motivo||"";
  rec.pendientes = rec.pendientes.filter(p=>!(p.auto && p.originEvId===evId));  // revierte auto-pendientes
  editing=null; saveHist();
}
function savePend(pId){
  const rec = ensureEq(CURRENT_EQ);
  const data = {titulo:val("f_pt"), tipo:val("f_ptipo")||"Manual", responsable:val("f_presp"), fechaLimite:val("f_pfl"), estado:val("f_pest")};
  if(!data.titulo){ editing=null; return renderDrawer(); }
  if(pId){ Object.assign(rec.pendientes.find(x=>x.id===pId), data); }
  else { rec.pendientes.push({id:uid(), ...data, auto:false, folio:"", tareas:[]}); }
  editing=null; saveHist();
}
function saveTask(pId,tId){
  const p = ensureEq(CURRENT_EQ).pendientes.find(x=>x.id===pId); if(!p) return;
  const data = {tipo:val("f_ttipo"), estado:val("f_test"), descripcion:val("f_tdesc"), nota:val("f_tnota")};
  if(!data.descripcion){ editing=null; return renderDrawer(); }
  p.tareas = p.tareas||[];
  if(tId){ Object.assign(p.tareas.find(x=>x.id===tId), data); } else { p.tareas.push({id:uid(), ...data}); }
  editing=null; saveHist();
}

const dBody = $("#dBody");
dBody.addEventListener("click", e=>{
  const b = e.target.closest("[data-act]"); if(!b) return;
  const {act, ev:evId, p:pId, t:tId} = b.dataset;
  if(act==="cancel"){ editing=null; return renderDrawer(); }
  if(act==="add-event"){ editing={kind:"event",mode:"add",tipo:ETIPOS[0]}; return renderDrawer(); }
  if(act==="edit-event"){ editing={kind:"event",mode:"edit",evId}; return renderDrawer(); }
  if(act==="save-event"){ return saveEvent(evId); }
  if(act==="anul-event"){ return anularEvento(evId); }
  if(act==="add-pend"){ editing={kind:"pend",mode:"add"}; return renderDrawer(); }
  if(act==="edit-pend"){ editing={kind:"pend",mode:"edit",pId}; return renderDrawer(); }
  if(act==="save-pend"){ return savePend(pId); }
  if(act==="del-pend"){ if(confirm("¿Eliminar este pendiente?")){ const rec=ensureEq(CURRENT_EQ); rec.pendientes=rec.pendientes.filter(x=>x.id!==pId); editing=null; saveHist(); } return; }
  if(act==="add-task"){ editing={kind:"task",mode:"add",pId}; return renderDrawer(); }
  if(act==="edit-task"){ editing={kind:"task",mode:"edit",pId,tId}; return renderDrawer(); }
  if(act==="save-task"){ return saveTask(pId,tId); }
  if(act==="del-task"){ const p=ensureEq(CURRENT_EQ).pendientes.find(x=>x.id===pId); p.tareas=p.tareas.filter(x=>x.id!==tId); editing=null; saveHist(); return; }
});
dBody.addEventListener("change", e=>{
  const t = e.target;
  if(t.id==="f_tipo" && editing && editing.kind==="event" && editing.mode==="add"){ editing.tipo=t.value; return renderDrawer(); }
  if(t.id==="f_resultado"){ const sp=document.getElementById("mpEst"); if(sp){ const ne=mpEstado(t.value); sp.textContent= ne?ne:"(sin cambio)"; } return; }
  if(t.id==="f_via"){ const w=document.getElementById("wrap_finforme"); if(w) w.style.display=(t.value==="Trato directo")?"":"none"; return; }
  const c = t.closest('[data-act="toggle-task"]'); if(!c) return;
  const p = ensureEq(CURRENT_EQ).pendientes.find(x=>x.id===c.dataset.p);
  const tk = p.tareas.find(x=>x.id===c.dataset.t);
  tk.estado = c.checked ? "hecha" : "pendiente"; saveHist();
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
        HIST = obj;
        eqIds().forEach(id=>{ HIST[id].eventos=HIST[id].eventos||[]; HIST[id].pendientes=HIST[id].pendientes||[]; });
        saveHist();
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
