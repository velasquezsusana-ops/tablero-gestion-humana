# -*- coding: utf-8 -*-
"""
Extrae los datos de "Gestión humana.xlsx", "Informe tipo de contrato.xlsx",
"INGRESOS.xlsx" y "ROTACION..xlsx" y genera gestion_humana_data.json, que
luego se inyecta en "Tablero Gestion Humana.html" (bloque `const DATA_RAW = {...}`
dentro del <script>), reemplazando la lectura en vivo por File System Access API.

Volver a correr este script (con C:\\Users\\Montolivo\\AppData\\Local\\Programs\\Python\\Python313\\python.exe)
cada vez que se actualicen los 4 excels, y luego pedirle a Claude que vuelva a
"splicear" gestion_humana_data.json dentro del HTML.
"""
import json
import os
import re
import unicodedata
import datetime
import openpyxl

BASE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.dirname(BASE)

MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
          "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
TRIMESTRES = ["Trimestre 1", "Trimestre 2", "Trimestre 3", "Trimestre 4"]


def mes_idx(m):
    try:
        return MESES.index(str(m).strip())
    except ValueError:
        return -1


def period_key(anio, mes):
    return int(anio) * 100 + mes_idx(mes)


def as_date(v):
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.isoformat()[:10]
    return None


def as_num(v):
    try:
        n = float(v)
        return n if n == n else 0  # NaN check
    except (TypeError, ValueError):
        return 0


def as_num_or_none(v):
    if v is None or v == '':
        return None
    try:
        n = float(v)
        return n if n == n else None
    except (TypeError, ValueError):
        return None


def find_sheet(wb, wanted_name):
    for n in wb.sheetnames:
        if n.strip() == wanted_name.strip():
            return wb[n]
    return None


def rows_from(ws, start_row):
    """All rows (as tuples of values) from start_row (1-based) to the end."""
    if ws is None:
        return []
    out = []
    for r in range(start_row, ws.max_row + 1):
        out.append([ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)])
    return out


def g(row, idx):
    return row[idx] if idx < len(row) else None


out = {
    "clima": [], "formacion": [], "costo": [], "costoContratacion": [],
    "sst": [], "ausentismoGH": [], "costoBonificacion": [], "cobroIncapacidades": [],
    "depuracionCartera": [], "nomina": [], "ingresos": [], "rotacion": [],
}

# ── Gestión humana.xlsx ──────────────────────────────────────────────
gh_file = next(f for f in os.listdir(DATA_DIR) if f.lower().startswith("gesti") and f.lower().endswith(".xlsx") and not f.startswith("~$"))
wb_gh = openpyxl.load_workbook(os.path.join(DATA_DIR, gh_file), data_only=True)

for r in rows_from(find_sheet(wb_gh, "Clima organizacional"), 2):
    if g(r, 1) is None or g(r, 3) is None:
        continue
    out["clima"].append({
        "anio": g(r, 1), "trimestre": g(r, 2), "area": g(r, 3),
        "respondientes": g(r, 4), "promedio": g(r, 5), "pct": g(r, 6),
        "objetivoPts": g(r, 7), "objetivoPct": g(r, 8),
    })

for r in rows_from(find_sheet(wb_gh, "Plan de formación"), 2):
    if g(r, 0) is None or g(r, 1) is None:
        continue
    out["formacion"].append({
        "anio": g(r, 0), "mes": g(r, 1), "planificadas": g(r, 2),
        "realizadas": g(r, 3), "cumplimiento": g(r, 4), "objetivo": g(r, 5),
        "observaciones": g(r, 7),
    })

# Costo de capacitación: Año,Mes,Papelería,Cursos,Traslados Internos,Locativos,
# Plan padrino PDV,Cursos de Alturas,Curso coord. Altura,Curso BPM,Total,Empleados,Costo/empleado
for r in rows_from(find_sheet(wb_gh, "Costo de capacitación"), 2):
    if g(r, 0) is None or g(r, 1) is None:
        continue
    out["costo"].append({
        "anio": g(r, 0), "mes": g(r, 1), "papeleria": as_num(g(r, 2)), "cursos": as_num(g(r, 3)),
        "trasladosInternos": as_num(g(r, 4)), "locativos": as_num(g(r, 5)),
        "planPadrinoPdv": as_num(g(r, 6)), "cursosAlturas": as_num(g(r, 7)),
        "cursoCoordAltura": as_num(g(r, 8)), "cursoBpm": as_num(g(r, 9)),
        "total": as_num(g(r, 10)), "empleados": g(r, 11), "costoEmpleado": as_num_or_none(g(r, 12)),
    })

# Costo de contratación (antes "Tiempo de contratación"): Año,Mes,Vacante,
# Costo examenes medicos Mujer,Costo examenes medicos Hombre,Estudio confiabilidad,Dotación,
# Total costo contratación,total costo mujer,total costo hombre,empleados mujer,empleados hombre,
# total empleados contratados,Costo/empleado  (columnas Mujer/Hombre nuevas en la hoja)
for r in rows_from(find_sheet(wb_gh, " Costo de contratación "), 3):
    if g(r, 0) is None or g(r, 1) is None or g(r, 2) is None:
        continue
    out["costoContratacion"].append({
        "anio": g(r, 0), "mes": g(r, 1), "vacante": g(r, 2),
        "examenesMedicosMujer": as_num_or_none(g(r, 3)), "examenesMedicosHombre": as_num_or_none(g(r, 4)),
        "estudioConfiabilidad": as_num_or_none(g(r, 5)), "dotacion": as_num_or_none(g(r, 6)),
        "total": as_num(g(r, 7)),
        "totalCostoMujer": as_num_or_none(g(r, 8)), "totalCostoHombre": as_num_or_none(g(r, 9)),
        "empleadosMujer": as_num_or_none(g(r, 10)), "empleadosHombre": as_num_or_none(g(r, 11)),
        "empleadosContratados": g(r, 12), "costoEmpleado": as_num_or_none(g(r, 13)),
    })

# SST (antes "Accidentes laborales"): registro detallado por accidente.
# Año,Mes,Fecha,Día del accidente,Cedula,Nombre,Cargo,Proceso,Punto de venta,Accidente,
# Causal,tipo de accidente,Severidad  (Proceso/Punto de venta/tipo de accidente son columnas nuevas)
for r in rows_from(find_sheet(wb_gh, "SST"), 2):
    if g(r, 0) is None and g(r, 1) is None and g(r, 5) is None:
        continue
    out["sst"].append({
        "anio": g(r, 0), "mes": g(r, 1), "fecha": as_date(g(r, 2)), "diaAccidente": g(r, 3),
        "cedula": g(r, 4), "nombre": g(r, 5), "cargo": g(r, 6),
        "proceso": g(r, 7), "puntoVenta": g(r, 8), "accidente": g(r, 9),
        "causal": g(r, 10), "tipoAccidente": g(r, 11), "severidad": g(r, 12),
    })

# Ausentismo (hoja nueva): Año,Mes,Concepto,Días,Valor($)
for r in rows_from(find_sheet(wb_gh, "Ausentismo"), 2):
    if g(r, 0) is None and g(r, 1) is None and g(r, 2) is None:
        continue
    out["ausentismoGH"].append({
        "anio": g(r, 0), "mes": g(r, 1), "concepto": g(r, 2),
        "dias": as_num_or_none(g(r, 3)), "valor": as_num_or_none(g(r, 4)),
    })

# Costo de Bonificación (hoja nueva, reestructurada a formato largo por concepto):
# Año,Mes,Concepto,valor,factor prestacional,total costo,nota
for r in rows_from(find_sheet(wb_gh, "Costo de Bonificación"), 2):
    if g(r, 0) is None and g(r, 1) is None and g(r, 2) is None:
        continue
    out["costoBonificacion"].append({
        "anio": g(r, 0), "mes": g(r, 1), "concepto": g(r, 2),
        "valor": as_num_or_none(g(r, 3)), "factorPrestacional": as_num_or_none(g(r, 4)),
        "totalCosto": as_num_or_none(g(r, 5)), "nota": g(r, 6),
    })

# Cobro de incapacidades (hoja nueva, reestructurada a detalle por caso de incapacidad):
# Mes,Año,Cédula,Nombre,Tipo de Novedad,Fecha inicio,Fecha fin,Días de novedad,Eps,
# Radicado/Trámite,Gestiona S/N,Fecha de radicado,Observación,Salario,Días a cobrar,
# Valor a recuperar,Valor pagado,Fecha de pago,Estado del pago
for r in rows_from(find_sheet(wb_gh, "Cobro de incapacidades"), 2):
    if g(r, 0) is None and g(r, 1) is None and g(r, 2) is None:
        continue
    out["cobroIncapacidades"].append({
        "mes": g(r, 0), "anio": g(r, 1), "cedula": g(r, 2), "nombre": g(r, 3),
        "tipoNovedad": g(r, 4), "fechaInicio": as_date(g(r, 5)), "fechaFin": as_date(g(r, 6)),
        "diasNovedad": as_num_or_none(g(r, 7)), "administradoraSalud": g(r, 8),
        "radicado": g(r, 9), "gestiona": g(r, 10), "fechaRadicado": as_date(g(r, 11)),
        "observacion": g(r, 12), "salario": as_num_or_none(g(r, 13)),
        "diasACobrar": as_num_or_none(g(r, 14)),
        "valorCobradas": as_num_or_none(g(r, 15)), "valorPagadas": as_num_or_none(g(r, 16)),
        "fechaPago": as_date(g(r, 17)), "estadoPago": g(r, 18),
    })

# Depuración de Cartera (hoja nueva): 4 bloques (pensión / cajas de compensación /
# salud / ARL), cada uno con encabezado "Año | Mes | <categoría>" y filas de entidad debajo.
ws_dep = find_sheet(wb_gh, "Depuración de Cartera")
if ws_dep is not None:
    current_categoria = None
    for r in range(1, ws_dep.max_row + 1):
        row = [ws_dep.cell(row=r, column=c).value for c in range(1, ws_dep.max_column + 1)]
        a, b, c = g(row, 0), g(row, 1), g(row, 2)
        f_, gg, h, i = g(row, 3), g(row, 4), g(row, 5), g(row, 6)
        if a is not None and str(a).strip().lower().startswith("año"):
            current_categoria = c
            continue
        if c is not None:
            out["depuracionCartera"].append({
                "categoria": current_categoria, "entidad": c, "anio": a, "mes": b,
                "estadoCuenta": f_, "valorCorreccion": as_num_or_none(gg),
                "valorAPagar": as_num_or_none(h), "valorDevolucion": as_num_or_none(i),
            })

out["formacion"].sort(key=lambda x: period_key(x["anio"], x["mes"]))
out["costo"].sort(key=lambda x: period_key(x["anio"], x["mes"]))
out["clima"].sort(key=lambda x: (int(x["anio"]), TRIMESTRES.index(x["trimestre"]) if x["trimestre"] in TRIMESTRES else -1))

# ── Informe tipo de contrato.xlsx ────────────────────────────────────
# Reporte histórico de contratos (varias filas por empleado = renovaciones).
# "Nómina y contratación" solo necesita el contrato más reciente de cada empleado.
wb_ct = openpyxl.load_workbook(os.path.join(DATA_DIR, "Informe tipo de contrato.xlsx"), data_only=True)
ws_ct = wb_ct[wb_ct.sheetnames[0]]
latest_by_id = {}
for r in rows_from(ws_ct, 2):
    emp = g(r, 0)
    if emp is None:
        continue
    fecha_ingreso = g(r, 5)
    if not isinstance(fecha_ingreso, (datetime.datetime, datetime.date)):
        continue
    prev = latest_by_id.get(emp)
    if prev is None or fecha_ingreso > prev["_fi"]:
        latest_by_id[emp] = {
            "_fi": fecha_ingreso,
            "id": emp, "nombre": g(r, 1), "cargo": g(r, 2), "co": g(r, 3), "centroCosto": g(r, 4),
            "fechaIngreso": as_date(fecha_ingreso), "fechaContratoHasta": as_date(g(r, 6)),
        }
for rec in latest_by_id.values():
    del rec["_fi"]
    out["nomina"].append(rec)
out["nomina"].sort(key=lambda x: x["nombre"] or "")

# ── INGRESOS.xlsx ─────────────────────────────────────────────────────
CITY_SHEETS = ["MEDELLÍN-RIONEGRO", "RIONEGRO", "BOGOTA", "BARRANQUILLA", "EJE CAFETERO"]
OTHER_SHEETS = {"NO CRITICOS": "no_criticos", "TEMPORADA": "temporada", "APRENDICES ": "aprendices"}
wb_ing = openpyxl.load_workbook(os.path.join(DATA_DIR, "INGRESOS.xlsx"), data_only=True)

def extract_ingresos(ws, ciudad, tipo):
    for r in rows_from(ws, 2):
        nombre, cargo, estado = g(r, 0), g(r, 1), g(r, 2)
        if nombre is None and cargo is None and estado is None:
            continue
        fecha_ingreso = g(r, 3)
        fecha_solicitud = g(r, 5)
        dias = None
        if isinstance(fecha_ingreso, (datetime.datetime, datetime.date)) and isinstance(fecha_solicitud, (datetime.datetime, datetime.date)):
            dias = (fecha_ingreso - fecha_solicitud).days
        estado_norm = str(estado).strip().upper() if estado else None
        out["ingresos"].append({
            "ciudad": ciudad, "tipo": tipo, "nombre": nombre,
            "cargo": str(cargo).strip().upper() if cargo else None,
            "estado": estado_norm, "contratado": bool(estado_norm and "CONTRATADO" in estado_norm),
            "fechaIngreso": as_date(fecha_ingreso), "fechaSolicitud": as_date(fecha_solicitud),
            "dias": dias, "fechaMaxima": as_date(g(r, 10)), "estadoIngreso": g(r, 12),
        })

for sn in CITY_SHEETS:
    extract_ingresos(wb_ing[sn], sn, "ciudad")
for sn, tipo in OTHER_SHEETS.items():
    extract_ingresos(wb_ing[sn], sn.strip(), tipo)

# ── ROTACION..xlsx ────────────────────────────────────────────────────
MES_ALIASES = {
    "ENERO": "Enero", "ENE": "Enero", "FEBRERO": "Febrero", "FEB": "Febrero",
    "MARZO": "Marzo", "MAR": "Marzo", "ABRIL": "Abril", "ABR": "Abril",
    "MAYO": "Mayo", "JUNIO": "Junio", "JULIO": "Julio",
    "AGOSTO": "Agosto", "AGO": "Agosto", "SEPTIEMBRE": "Septiembre", "SEP": "Septiembre",
    "OCTUBRE": "Octubre", "OCT": "Octubre", "NOVIEMBRE": "Noviembre", "NOV": "Noviembre",
    "DICIEMBRE": "Diciembre", "DIC": "Diciembre",
}
ROT_LABELS = {
    "VINCULADOS": "vinculados", "TEMPORALES": "temporales", "TOTAL PERSONAL": "totalPersonal",
    "RETIROS": "totalRetiros", "TOTAL RETIROS": "totalRetiros", "TOTAL ROTACION": "totalRotacion",
}


def strip_accents(s):
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')


def norm_label(s):
    return strip_accents(str(s)).upper().strip()


def parse_rotacion_sheet_name(name, carry_year):
    n = norm_label(name)
    n = re.sub(r'^ROTACION\.?\s*', '', n)
    n = re.sub(r'\(\d+\)\s*$', '', n).strip()
    ym = re.search(r'20\d{2}', n)
    year = int(ym.group(0)) if ym else carry_year
    rest = (n[:ym.start()] if ym else n).strip()
    mes = None
    for alias in sorted(MES_ALIASES.keys(), key=len, reverse=True):
        if rest.startswith(alias):
            mes = MES_ALIASES[alias]
            break
    return year, mes


def extract_rotacion_metrics(ws):
    res = {}
    max_c = min(ws.max_column, 12)
    for r in range(1, 10):
        for c in range(1, max_c + 1):
            v = ws.cell(row=r, column=c).value
            if not isinstance(v, str):
                continue
            lbl = norm_label(v)
            if lbl in ROT_LABELS:
                key = ROT_LABELS[lbl]
                if key in res:
                    continue
                val = ws.cell(row=r, column=c + 1).value
                if isinstance(val, (int, float)):
                    res[key] = val
    return res


wb_rot = openpyxl.load_workbook(os.path.join(DATA_DIR, "ROTACION..xlsx"), data_only=True)
carry_year = None
seen_periods = {}
for sn in wb_rot.sheetnames:
    year, mes = parse_rotacion_sheet_name(sn, carry_year)
    carry_year = year
    if mes is None:
        continue
    key = (year, mes)
    if key in seen_periods:
        continue  # duplicate sheet (e.g. "(2)"), same data
    metrics = extract_rotacion_metrics(wb_rot[sn])
    if not metrics:
        continue
    seen_periods[key] = True
    out["rotacion"].append({"anio": year, "mes": mes, **metrics})
out["rotacion"].sort(key=lambda x: period_key(x["anio"], x["mes"]))

out_path = os.path.join(BASE, "gestion_humana_data.json")
with open(out_path, "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, indent=2)

print("OK ->", out_path)
for k, v in out.items():
    print(f"  {k}: {len(v)} filas")
