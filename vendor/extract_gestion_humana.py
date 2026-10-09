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


def parse_money(v):
    """Convierte a numero tanto celdas numericas normales como celdas de texto con
    formato de moneda escritas a mano (p.ej. ' $      102.850' con '.' como separador
    de miles). Errores de formula (#DIV/0!, #N/A...) devuelven None."""
    if v is None or v == '':
        return None
    if isinstance(v, (int, float)):
        return v if v == v else None  # descarta NaN
    s = str(v).strip()
    if not s or s.startswith('#'):
        return None
    s = s.replace('$', '').replace('\xa0', ' ').strip()
    if s == '-':
        return 0  # notación contable de "cero"
    s = s.replace('.', '').replace(',', '.')
    try:
        n = float(s)
        return n if n == n else None
    except ValueError:
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


def strip_accents(s):
    return ''.join(c for c in unicodedata.normalize('NFD', str(s)) if unicodedata.category(c) != 'Mn')


out = {
    "clima": [], "formacion": [], "costo": [], "costoContratacion": [],
    "sst": [], "ausentismoGH": [], "costoBonificacion": [], "cobroIncapacidades": [],
    "depuracionCartera": [], "nomina": [], "ingresos": [], "rotacion": [],
    "rotacionCausa": [], "rotacionCiudad": [], "capacitacionesDetalle": [],
    "carteraInconsistencias": [],
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

# Detalle de capacitaciones (archivos "CAPACITACIONES <MES> <AÑO>.xlsx", uno por mes,
# sueltos en la misma carpeta). Cada uno trae 1 hoja: Capacitación,Fecha,Encargada,Lugar,Asistencia.
capacitaciones_files = [f for f in os.listdir(DATA_DIR)
                         if f.upper().startswith("CAPACITACIONES") and f.lower().endswith(".xlsx") and not f.startswith("~$")]
for cf in capacitaciones_files:
    wb_cap = openpyxl.load_workbook(os.path.join(DATA_DIR, cf), data_only=True)
    for sheetname in wb_cap.sheetnames:
        for r in rows_from(wb_cap[sheetname], 2):
            nombre_cap = g(r, 0)
            if nombre_cap is None:
                continue
            fecha_cap = g(r, 1)
            anio_cap = fecha_cap.year if isinstance(fecha_cap, (datetime.datetime, datetime.date)) else None
            mes_cap = MESES[fecha_cap.month - 1] if isinstance(fecha_cap, (datetime.datetime, datetime.date)) else None
            out["capacitacionesDetalle"].append({
                "anio": anio_cap, "mes": mes_cap, "capacitacion": str(nombre_cap).strip(),
                "fecha": as_date(fecha_cap), "encargada": g(r, 2), "lugar": g(r, 3), "asistencia": g(r, 4),
            })
out["capacitacionesDetalle"].sort(key=lambda x: x["fecha"] or "")

# Costo de capacitación: las columnas se buscan por ENCABEZADO (no por posición), porque
# la hoja va ganando columnas nuevas (sept-2026: Plan Padrino Cocina / Servicio, personas
# entrenadas, valor unitario y # personas de curso de alturas). Filas sin "Total
# capacitación" (plantilla de meses que aún no cierran) se descartan.
COSTO_COLS = [  # (campo, prefijo del encabezado normalizado)
    ("papeleria", "PAPELERIA"), ("cursos", "CURSOS ($"), ("trasladosInternos", "TRASLADOS"),
    ("locativos", "LOCATIVOS"), ("planPadrinoPdv", "PLAN PADRINO PDV"),
    ("planPadrinoCocina", "PLAN PADRINO COCINA"), ("entrenadosCocina", "NUMERO DE PERSONAS ENTRENADAS COCINA"),
    ("planPadrinoServicio", "PLAN PADRINO SERVICIO"), ("entrenadosServicio", "NUMERO DE PERSONAS ENTRENADAS SERVICIO"),
    ("cursosAlturas", "CURSOS DE ALTURAS"), ("cursoCoordAltura", "CURSO COORD"), ("cursoBpm", "CURSO BPM"),
    ("total", "TOTAL"), ("empleados", "EMPLEADOS"), ("costoEmpleado", "COSTO POR EMPLEADO"),
]
ws_costo = find_sheet(wb_gh, "Costo de capacitación")
if ws_costo is not None:
    head = [strip_accents(ws_costo.cell(row=1, column=c).value or "").upper().strip() for c in range(1, ws_costo.max_column + 1)]
    cidx = {campo: next((k for k, h in enumerate(head) if h.startswith(pref)), None) for campo, pref in COSTO_COLS}
    for r in rows_from(ws_costo, 2):
        if g(r, 0) is None or g(r, 1) is None:
            continue
        val = lambda campo: g(r, cidx[campo]) if cidx[campo] is not None else None
        if not as_num(val("total")):
            continue
        rec = {"anio": g(r, 0), "mes": g(r, 1)}
        for campo, _ in COSTO_COLS:
            if campo == "empleados":
                rec[campo] = val(campo)
            elif campo == "costoEmpleado":
                rec[campo] = as_num_or_none(val(campo))
            else:
                rec[campo] = as_num(val(campo))
        out["costo"].append(rec)

# Costo de contratación (antes "Tiempo de contratación"): Año,Mes,Vacante,
# Costo examenes medicos Mujer,Costo examenes medicos Hombre,Estudio confiabilidad,Dotación,
# Curso BPM,Valor unitario presencial,Cantidad presencial,Valor unitario virtual,Cantidad virtual,
# Total costo contratación,total costo mujer,total costo hombre,empleados mujer,empleados hombre,
# total empleados contratados,Costo/empleado
# (Curso BPM/Presencial/Virtual son columnas nuevas; la hoja ahora trae una fila plantilla
# por cada mes del año aunque no tenga datos reales -> se descartan filas sin empleados contratados)
for r in rows_from(find_sheet(wb_gh, " Costo de contratación "), 3):
    if g(r, 0) is None or g(r, 1) is None or g(r, 2) is None:
        continue
    if not g(r, 17):  # sin "Total empleados contratados" (None o 0) = fila plantilla sin datos reales
        continue
    out["costoContratacion"].append({
        "anio": g(r, 0), "mes": g(r, 1), "vacante": g(r, 2),
        "examenesMedicosMujer": parse_money(g(r, 3)), "examenesMedicosHombre": parse_money(g(r, 4)),
        "estudioConfiabilidad": parse_money(g(r, 5)), "dotacion": parse_money(g(r, 6)),
        "cursoBpm": parse_money(g(r, 7)),
        "valorUnitarioPresencial": parse_money(g(r, 8)), "cantidadPresencial": as_num_or_none(g(r, 9)),
        "valorUnitarioVirtual": parse_money(g(r, 10)), "cantidadVirtual": as_num_or_none(g(r, 11)),
        "total": parse_money(g(r, 12)) or 0,
        "totalCostoMujer": parse_money(g(r, 13)), "totalCostoHombre": parse_money(g(r, 14)),
        "empleadosMujer": as_num_or_none(g(r, 15)), "empleadosHombre": as_num_or_none(g(r, 16)),
        "empleadosContratados": g(r, 17), "costoEmpleado": parse_money(g(r, 18)),
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

# Ausentismo: desde sept-2026 la hoja es una TABLA DINÁMICA (fila de meses "jun/jul/ago/sep",
# debajo "Etiquetas de fila | Suma de Dias | Suma de Devengo" por mes, una fila por concepto,
# hasta "Total general"). No trae año: se toma el año más reciente del Plan de formación.
# Si la hoja vuelve al formato largo (Año,Mes,Concepto,Días,Valor) se lee como antes.
MES_CORTO = {m[:3].upper(): m for m in MESES}
ws_aus = find_sheet(wb_gh, "Ausentismo")
if ws_aus is not None:
    aus_rows = rows_from(ws_aus, 1)
    hdr_i = next((i for i, r in enumerate(aus_rows) if str(g(r, 0) or "").strip().lower().startswith("etiquetas de fila")), None)
    if hdr_i is not None:
        anio_aus = max([int(x["anio"]) for x in out["formacion"] if x.get("anio")] or [datetime.date.today().year])
        mes_row = next((aus_rows[i] for i in range(hdr_i - 1, -1, -1)
                        if any(strip_accents(str(v or "")).upper().strip()[:3] in MES_CORTO for v in aus_rows[i])), [])
        cols = []  # (mes, col días, col valor)
        for c, v in enumerate(mes_row):
            m = MES_CORTO.get(strip_accents(str(v or "")).upper().strip()[:3])
            if m and "DIAS" in strip_accents(str(g(aus_rows[hdr_i], c) or "")).upper():
                cols.append((m, c, c + 1))
        for r in aus_rows[hdr_i + 1:]:
            concepto = g(r, 0)
            if concepto is None or str(concepto).strip().lower().startswith("total"):
                continue
            for m, cd, cv in cols:
                dias, valor = as_num_or_none(g(r, cd)), as_num_or_none(g(r, cv))
                if dias is None and valor is None:
                    continue
                out["ausentismoGH"].append({
                    "anio": anio_aus, "mes": m, "concepto": str(concepto).strip(),
                    "dias": round(dias, 2) if dias is not None else None, "valor": valor,
                })
    else:
        for r in rows_from(ws_aus, 2):
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
        # Tablita "Entidades | Valor de inconsistencia" (sept-2026): una fila por entidad,
        # sin año/mes, hasta la fila TOTAL. Va aparte en carteraInconsistencias.
        if a is not None and str(a).strip().lower().startswith("entidades") and "inconsistencia" in strip_accents(str(b or "")).lower():
            current_categoria = "__inconsistencias__"
            continue
        if current_categoria == "__inconsistencias__":
            if a is not None and not str(a).strip().lower().startswith("total") and as_num_or_none(b) is not None:
                out["carteraInconsistencias"].append({"entidad": str(a).strip(), "valor": as_num_or_none(b)})
            if a is not None and str(a).strip().lower().startswith("año"):
                current_categoria = c
            continue
        if a is not None and str(a).strip().lower().startswith("año"):
            current_categoria = c
            continue
        if c is not None:
            out["depuracionCartera"].append({
                "categoria": current_categoria, "entidad": c, "anio": a, "mes": b,
                "estadoCuenta": f_, "valorCorreccion": as_num_or_none(gg),
                "valorAPagar": as_num_or_none(h), "valorDevolucion": as_num_or_none(i),
            })

# Rotación de personal — causas (hoja nueva): matriz Causa x Mes(2026) con
# encabezados de fecha en la fila 3; cada celda no vacía = # retiros de esa causa ese mes.
ws_rotcausa = find_sheet(wb_gh, "Rotacion de personal")
if ws_rotcausa is not None:
    header_row = [ws_rotcausa.cell(row=3, column=c).value for c in range(1, ws_rotcausa.max_column + 1)]
    month_cols = [(i, v.year, MESES[v.month - 1]) for i, v in enumerate(header_row)
                  if isinstance(v, (datetime.datetime, datetime.date))]
    r = 4
    while True:
        causa = ws_rotcausa.cell(row=r, column=1).value
        if causa is None or str(causa).strip().lower().startswith("total"):
            break
        for (idx, anio_c, mes_c) in month_cols:
            val = ws_rotcausa.cell(row=r, column=idx + 1).value
            if isinstance(val, (int, float)) and val != 0:
                out["rotacionCausa"].append({"anio": anio_c, "mes": mes_c, "causa": str(causa).strip(), "cantidad": val})
        r += 1
out["rotacionCausa"].sort(key=lambda x: period_key(x["anio"], x["mes"]))

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
# Hojas no-ciudad (utilitarias) a ignorar. Las hojas de ciudad no se listan por
# nombre fijo (ese nombre ha cambiado, p.ej. "MEDELLÍN-RIONEGRO" -> "MEDELLÍN"):
# cualquier hoja que no esté en NON_CITY_SHEETS ni en OTHER_SHEETS se trata como ciudad.
NON_CITY_SHEETS = {"FESTIVOS", "CARGOS Y ESTADOS", "HOJA1"}
OTHER_SHEETS = {"NO CRITICOS": "no_criticos", "TEMPORADA": "temporada", "APRENDICES": "aprendices"}
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

for sn in wb_ing.sheetnames:
    key = strip_accents(sn).strip().upper()
    if key in NON_CITY_SHEETS:
        continue
    if key in OTHER_SHEETS:
        extract_ingresos(wb_ing[sn], sn.strip(), OTHER_SHEETS[key])
    else:
        extract_ingresos(wb_ing[sn], sn.strip(), "ciudad")

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


CIUDADES_ROT = ["MEDELLIN", "BOGOTA", "PEREIRA", "BARRANQUILLA", "CALI", "MANIZALES"]


def extract_rotacion_ciudad(ws):
    """Recorre los bloques rol×ciudad (7 columnas repetidas: SERVICIO/CAJEROS/AUXILIAR
    DE COCINA/LIDERES/LIDER JUNIOR, cada uno con sub-bloques por ciudad) y agrega
    personal/retiros por ciudad. Los roles centralizados (producción, logística,
    mantenimiento, administrativos, servicios generales) no traen ciudad en la hoja
    fuente, así que quedan fuera de este desglose (es un desglose parcial, no total)."""
    max_c, max_r = ws.max_column, ws.max_row
    header_cols = [1, 4, 7, 10, 13, 16, 19]
    agg = {}
    for r in range(1, max_r + 1):
        for c in header_cols:
            if c > max_c:
                continue
            v = ws.cell(row=r, column=c).value
            if not isinstance(v, str):
                continue
            nv = norm_label(v)
            if any(k in nv for k in ('RETIRO', 'TOTAL', 'INCULAD', 'ROTACI', 'TEMPORAL')):
                continue  # no es un encabezado de bloque, es una fila de datos
            ciudad = next((ct for ct in CIUDADES_ROT if ct in nv), None)
            if not ciudad:
                continue
            vinc = temp = ret = None
            for rr in range(r + 1, min(r + 7, max_r + 1)):
                lbl = ws.cell(row=rr, column=c).value
                val = ws.cell(row=rr, column=c + 1).value
                if lbl is None:
                    continue
                nl = norm_label(lbl)
                if 'INCULAD' in nl and vinc is None and isinstance(val, (int, float)):
                    vinc = val
                elif 'TEMPORAL' in nl and isinstance(val, (int, float)):
                    temp = val
                elif 'RETIRO' in nl and isinstance(val, (int, float)):
                    ret = val
                elif 'ROTACI' in nl and isinstance(val, (int, float)):
                    break
            if vinc is not None and ret is not None:
                e = agg.setdefault(ciudad, {"personal": 0, "retiros": 0})
                e["personal"] += vinc + (temp or 0)
                e["retiros"] += ret
    return agg


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
    ciudad_agg = extract_rotacion_ciudad(wb_rot[sn])
    for ciudad, e in ciudad_agg.items():
        pct = (e["retiros"] / e["personal"] * 100) if e["personal"] else 0
        out["rotacionCiudad"].append({
            "anio": year, "mes": mes, "ciudad": ciudad,
            "personal": e["personal"], "retiros": e["retiros"], "rotacionPct": pct,
        })
out["rotacion"].sort(key=lambda x: period_key(x["anio"], x["mes"]))
out["rotacionCiudad"].sort(key=lambda x: period_key(x["anio"], x["mes"]))

# En los Excel a veces el mes viene con espacios ("Septiembre "): se limpian para que
# ordene, filtre y cruce bien con el detalle de CAPACITACIONES.
for _lista in out.values():
    for _r in _lista:
        if isinstance(_r.get("mes"), str):
            _r["mes"] = _r["mes"].strip()
out["formacion"].sort(key=lambda x: period_key(x["anio"], x["mes"]))
out["costo"].sort(key=lambda x: period_key(x["anio"], x["mes"]))

out_path = os.path.join(BASE, "gestion_humana_data.json")
with open(out_path, "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, indent=2)

print("OK ->", out_path)
for k, v in out.items():
    print(f"  {k}: {len(v)} filas")
