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


def _norm_hoja(n):
    return " ".join(strip_accents(str(n)).lower().split())


def find_sheet(wb, wanted_name, prefijo=None):
    """Busca la hoja sin importar tildes, mayúsculas ni espacios de más. Si no aparece y se da
    `prefijo`, sirve la primera hoja cuyo nombre empiece así (p.ej. "Rotación de pesonal ")."""
    for n in wb.sheetnames:
        if _norm_hoja(n) == _norm_hoja(wanted_name):
            return wb[n]
    if prefijo:
        for n in wb.sheetnames:
            if _norm_hoja(n).startswith(_norm_hoja(prefijo)):
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
    "carteraInconsistencias": [], "aprendices": [], "costoContratacionMensual": [],
}

# ── Gestión humana.xlsx ──────────────────────────────────────────────
# Los archivos se buscan por el comienzo del nombre ("ROTACION..xlsx" o "ROTACION.xlsx" sirven igual),
# ignorando las copias viejas que se guardan al lado ("... ultima version", "respaldo", "backup", "copia").
def es_respaldo(f):
    n = strip_accents(f).lower()
    return f.startswith("~$") or any(w in n for w in ("ultima", "ulrima", "respaldo", "backup", "copia"))

def buscar_xlsx(prefijo, obligatorio=True):
    cands = [f for f in os.listdir(DATA_DIR)
             if strip_accents(f).lower().startswith(prefijo) and f.lower().endswith(".xlsx") and not es_respaldo(f)]
    if not cands:
        if obligatorio:
            raise SystemExit(f'No encontré el archivo "{prefijo}*.xlsx" en {DATA_DIR}')
        return None
    return max(cands, key=lambda f: os.path.getmtime(os.path.join(DATA_DIR, f)))

gh_file = buscar_xlsx("gesti")
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
        # oct-2026: una hoja por mes y columna nueva "NUMERO DE ASISTENTES" (se busca por encabezado)
        head_cap = [strip_accents(str(c.value or "")).upper().strip() for c in wb_cap[sheetname][1]]
        c_asis = next((k for k, h in enumerate(head_cap) if h.startswith("NUMERO DE ASISTENTES")), None)
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
                "asistentes": as_num_or_none(g(r, c_asis)) if c_asis is not None else None,
                # Inducción = ingresos (no cuenta como planeada); lo demás es capacitación
                "tipo": "Inducción" if "INDUCCION" in strip_accents(str(nombre_cap)).upper() else "Capacitación",
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
# Oct-2026: la hoja se dividió en dos bloques. Fila 1 = títulos de bloque
# ("COSTOS DE CONTRATACIÓN (FORMACIÓN, BIENESTAR Y DESARROLLO)" | "COSTOS DE FORMACIÓN CONSTANTE
# Y BIENESTAR") y los encabezados bajaron a la fila 2. Los campos de siempre se buscan en el
# bloque 1; los del bloque 2 van con prefijo "fc". Si la hoja vuelve a un solo bloque, igual sirve.
FC_COLS = [
    ("fcCursoCoordAltura", "CURSO COORD"), ("fcCursosAlturas", "CURSOS DE ALTURAS"), ("fcCursoBpm", "CURSO BPM"),
    ("fcCapacitacionesGenerales", "COSTOS DE CAPACITACIONES GENERALES"), ("fcBienestar", "COSTOS DE ACTIVIDADES DE BIENESTAR"),
]
ws_costo = find_sheet(wb_gh, "Costo de capacitación")
if ws_costo is not None:
    def _fila_norm(rn):
        return [strip_accents(str(ws_costo.cell(row=rn, column=c).value or "")).upper().strip() for c in range(1, ws_costo.max_column + 1)]
    hr = 1 if _fila_norm(1)[:1] == ["ANO"] else 2
    head = _fila_norm(hr)
    titulos = _fila_norm(1) if hr == 2 else []
    b2 = next((k for k, t in enumerate(titulos) if "FORMACION CONSTANTE" in t), len(head))
    def _buscar(pref, lo=0, hi=None):
        hi = len(head) if hi is None else hi
        return next((k for k in range(lo, hi) if head[k].startswith(pref)), None)
    cidx = {campo: _buscar(pref, 0, b2) for campo, pref in COSTO_COLS}
    cidx["empleados"] = _buscar("EMPLEADOS")
    cidx["costoEmpleado"] = _buscar("COSTO POR EMPLEADO")
    fc_fin = cidx["empleados"] if cidx["empleados"] is not None and cidx["empleados"] > b2 else None
    fcidx = {campo: _buscar(pref, b2, fc_fin) for campo, pref in FC_COLS}
    for r in rows_from(ws_costo, hr + 1):
        if g(r, 0) is None or g(r, 1) is None:
            continue
        val = lambda idx: g(r, idx) if idx is not None else None
        total1 = as_num(val(cidx["total"]))
        fc = {campo: as_num(val(fcidx[campo])) for campo, _ in FC_COLS}
        total2 = sum(fc.values())
        if not total1 and not total2:   # mes que aún no cierra
            continue
        rec = {"anio": g(r, 0), "mes": g(r, 1)}
        for campo, _ in COSTO_COLS:
            if campo == "empleados":
                rec[campo] = val(cidx[campo])
            elif campo == "costoEmpleado":
                rec[campo] = as_num_or_none(val(cidx[campo]))
            else:
                rec[campo] = as_num(val(cidx[campo]))
        rec.update(fc)
        rec["totalFormacionConstante"] = total2
        rec["totalGeneral"] = total1 + total2
        out["costo"].append(rec)

# Costo de contratación (antes "Tiempo de contratación"): Año,Mes,Vacante,
# Costo examenes medicos Mujer,Costo examenes medicos Hombre,Estudio confiabilidad,Dotación,
# Curso BPM,Valor unitario presencial,Cantidad presencial,Valor unitario virtual,Cantidad virtual,
# Total costo contratación,total costo mujer,total costo hombre,empleados mujer,empleados hombre,
# total empleados contratados,Costo/empleado
# (Curso BPM/Presencial/Virtual son columnas nuevas; la hoja ahora trae una fila plantilla
# por cada mes del año aunque no tenga datos reales -> se descartan filas sin empleados contratados)
# Oct-2026: columnas por ENCABEZADO (se insertaron columnas vacías entre los exámenes médicos).
# A la derecha viene una tabla resumen por mes: Año | Mes | Costos totales de contratación |
# Cantidad de empleados ingresados | Costo por empleado  -> costoContratacionMensual.
CC_COLS = [  # (campo, prefijo del encabezado normalizado, posición antigua por si no hay encabezado)
    ("examenesMedicosMujer", "COSTO EXAMENES MEDICOS MUJER", 3), ("examenesMedicosHombre", "COSTO EXAMENES MEDICOS HOMBRE", 4),
    ("estudioConfiabilidad", "ESTUDIO CONFIABILIDAD", 5), ("dotacion", "DOTACION", 6), ("cursoBpm", "CURSO BPM", 7),
    ("valorUnitarioPresencial", "VALOR UNITARIO PRESENCIAL", 8), ("cantidadPresencial", "CANTIDAD PRESENCIAL", 9),
    ("valorUnitarioVirtual", "VALOR UNITARIO VIRTUAL", 10), ("cantidadVirtual", "CANTIDAD VIRTUAL", 11),
    ("total", "TOTAL COSTO CONTRATACION", 12), ("totalCostoMujer", "TOTAL COSTO CONTRA MUJER", 13),
    ("totalCostoHombre", "TOTAL COSTO CONTRA HOMBRE", 14), ("empleadosMujer", "EMPLEADOS MUJER", 15),
    ("empleadosHombre", "EMPLEADOS HOMBRE", 16), ("empleadosContratados", "TOTAL EMPLEADOS CONTRATADOS", 17),
    ("costoEmpleado", "COSTO POR EMPLEADO (", 18),
]
ws_cc = find_sheet(wb_gh, " Costo de contratación ")
if ws_cc is not None:
    # encabezados en la fila donde la columna A dice "Año" (hoy es la fila 2)
    hr_cc = next((rn for rn in range(1, 6) if strip_accents(str(ws_cc.cell(row=rn, column=1).value or "")).upper().strip() == "ANO"), 2)
    head = [strip_accents(str(ws_cc.cell(row=hr_cc, column=c).value or "")).upper().strip() for c in range(1, ws_cc.max_column + 1)]
    def _cc(pref, default=None, desde=0):
        return next((k for k in range(desde, len(head)) if head[k].startswith(pref)), default)
    ci = {campo: _cc(pref, pos) for campo, pref, pos in CC_COLS}
    for r in rows_from(ws_cc, hr_cc + 1):
        if not isinstance(g(r, 0), (int, float)) or g(r, 1) is None or g(r, 2) is None:
            continue
        if not g(r, ci["empleadosContratados"]):  # fila plantilla del mes sin datos reales
            continue
        rec = {"anio": g(r, 0), "mes": g(r, 1), "vacante": g(r, 2)}
        for campo, _, _ in CC_COLS:
            v = g(r, ci[campo])
            if campo == "empleadosContratados":
                rec[campo] = v
            elif campo.startswith("cantidad") or campo.startswith("empleados"):
                rec[campo] = as_num_or_none(v)
            else:
                rec[campo] = parse_money(v)
        rec["total"] = rec["total"] or 0
        out["costoContratacion"].append(rec)

    c_tot = _cc("COSTOS TOTALES DE CONTRATACION")
    if c_tot is not None and c_tot >= 2:
        c_n = _cc("CANTIDAD DE EMPLEADOS INGRESADOS", None, c_tot)
        c_ce = _cc("COSTO POR EMPLEADO", None, c_tot + 1)
        for r in rows_from(ws_cc, hr_cc + 1):
            anio, mes = g(r, c_tot - 2), g(r, c_tot - 1)
            costo = parse_money(g(r, c_tot))
            n = as_num_or_none(g(r, c_n)) if c_n is not None else None
            if not isinstance(anio, (int, float)) or mes is None or not n:
                continue
            ce = as_num_or_none(g(r, c_ce)) if c_ce is not None else None
            out["costoContratacionMensual"].append({
                "anio": anio, "mes": mes, "costoTotal": costo, "empleados": n,
                "costoEmpleado": ce if ce is not None else ((costo or 0) / n),
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
        # Formato lista (oct-2026): Año | Mes | Comcepto | Dias | Costo | Factor Prestacional
        # (+ a la derecha una tablita con los % del factor, que se ignora). Columnas por encabezado.
        head = [strip_accents(v or "").upper().strip() for v in (aus_rows[0] if aus_rows else [])]
        def col(pred):
            return next((k for k, h in enumerate(head) if pred(h)), None)
        c_anio, c_mes = col(lambda h: h.startswith("ANO")), col(lambda h: h == "MES")
        c_con = col(lambda h: "CEPTO" in h)
        c_dias = col(lambda h: h.startswith("DIA"))
        c_val = col(lambda h: h.startswith("COSTO") or h.startswith("VALOR"))
        c_fac = col(lambda h: h.startswith("FACTOR"))
        for r in aus_rows[1:]:
            anio, mes, con = g(r, c_anio), g(r, c_mes), g(r, c_con)
            if anio is None or mes is None or con is None:
                continue
            valor = as_num_or_none(g(r, c_val)) if c_val is not None else None
            factor = as_num_or_none(g(r, c_fac)) if c_fac is not None else None
            out["ausentismoGH"].append({
                "anio": anio, "mes": str(mes).strip().capitalize(), "concepto": str(con).strip(),
                "dias": as_num_or_none(g(r, c_dias)), "valor": valor, "factorPrestacional": factor,
                "costoTotal": (valor or 0) + (factor or 0) if (valor is not None or factor is not None) else None,
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
# Desde oct-2026 viene en su propio archivo "Incapacidades.xlsx" (mismas columnas); si existe,
# manda ese, si no se usa la hoja del Excel de Gestión humana.
inc_file = buscar_xlsx("incapacidad", obligatorio=False)
ws_inc = openpyxl.load_workbook(os.path.join(DATA_DIR, inc_file), data_only=True).worksheets[0] if inc_file else find_sheet(wb_gh, "Cobro de incapacidades")
for r in rows_from(ws_inc, 2):
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
ws_rotcausa = find_sheet(wb_gh, "Rotacion de personal", prefijo="Rotacion de pe")
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
wb_ing = openpyxl.load_workbook(os.path.join(DATA_DIR, buscar_xlsx("ingresos")), data_only=True)

# Oct-2026: se agregó la columna "Dias de contratación" -> las demás se corrieron; se leen por
# ENCABEZADO (con la posición antigua de respaldo). Vacante = tiene fecha de solicitud, no tiene
# fecha de ingreso y el estado no es CONTRATADO (los días abiertos se calculan en el tablero vs. hoy).
def _es_fecha(v):
    return isinstance(v, (datetime.datetime, datetime.date))

def extract_ingresos(ws, ciudad, tipo):
    head = [strip_accents(str(ws.cell(row=1, column=c).value or "")).upper().strip() for c in range(1, ws.max_column + 1)]
    col = lambda name, pos: next((k for k, h in enumerate(head) if h == name), pos)
    cN, cC, cE = col("NOMBRE NUEVO INGRESO", 0), col("CARGO", 1), col("ESTADO", 2)
    cI, cSo, cS = col("FECHA DE INGRESO", 3), col("SOLICITA", 4), col("FECHA DE SOLICITUD", 5)
    cP, cR, cPos = col("PUNTO", 6), col("PERSONA A QUIEN REEMPLAZARA", 7), col("POSICION CLAVES", 8)
    cF, cEI = col("FECHA MAXIMA", 10), col("ESTADO DEL INGRESO", 12)
    txt = lambda v: (str(v).strip() or None) if v is not None else None
    for r in rows_from(ws, 2):
        nombre, cargo, estado = g(r, cN), g(r, cC), g(r, cE)
        if nombre is None and cargo is None and estado is None:
            continue
        fecha_ingreso = g(r, cI)
        fecha_solicitud = g(r, cS)
        dias = None
        if _es_fecha(fecha_ingreso) and _es_fecha(fecha_solicitud):
            dias = (fecha_ingreso - fecha_solicitud).days
        estado_norm = str(estado).strip().upper() if estado else None
        # "CONTRATADO" (y sus variantes mal escritas: CONTRADADO, CONTRATDO, CONTRATA)
        contratado = bool(estado_norm and estado_norm.startswith("CONTRA"))
        estado_ing = g(r, cEI)
        out["ingresos"].append({
            "ciudad": ciudad, "tipo": tipo, "nombre": nombre,
            "cargo": str(cargo).strip().upper() if cargo else None,
            "estado": estado_norm, "contratado": contratado,
            "vacante": (not _es_fecha(fecha_ingreso)) and _es_fecha(fecha_solicitud) and not contratado,
            "fechaIngreso": as_date(fecha_ingreso), "fechaSolicitud": as_date(fecha_solicitud),
            "dias": dias, "fechaMaxima": as_date(g(r, cF)),
            "estadoIngreso": estado_ing if isinstance(estado_ing, str) else None,
            "punto": txt(g(r, cP)), "reemplaza": txt(g(r, cR)), "solicita": txt(g(r, cSo)), "posicion": txt(g(r, cPos)),
        })

for sn in wb_ing.sheetnames:
    key = strip_accents(sn).strip().upper()
    if key in NON_CITY_SHEETS:
        continue
    if key in OTHER_SHEETS:
        extract_ingresos(wb_ing[sn], sn.strip(), OTHER_SHEETS[key])
    else:
        extract_ingresos(wb_ing[sn], sn.strip(), "ciudad")

# ── INGRESOS.xlsx / hoja APRENDICES ───────────────────────────────────
# Columnas por encabezado (esta hoja trae "FECHA DE FINALIZACION" que las otras no tienen).
# La sede (CO) y el área (centro de costo) salen de la nómina, cruzando el nombre por
# palabras: en INGRESOS va "NOMBRES APELLIDOS" y en la nómina "APELLIDOS NOMBRES".
def _tokens(s):
    return set(strip_accents(s or "").upper().split())

def match_nomina(nombre):
    a = _tokens(nombre)
    if not a:
        return None
    cands = [r for r in out["nomina"] if "APRENDIZ" in strip_accents(r.get("cargo") or "").upper()] or out["nomina"]
    best, k = None, 0
    for r in cands:
        kk = len(a & _tokens(r.get("nombre")))
        if kk > k:
            best, k = r, kk
    return best if k >= 2 and k >= min(3, len(a) - 1) else None

for sn in wb_ing.sheetnames:
    if strip_accents(sn).strip().upper() != "APRENDICES":
        continue
    ws_ap = wb_ing[sn]
    head = [strip_accents(ws_ap.cell(row=1, column=c).value or "").upper().strip() for c in range(1, ws_ap.max_column + 1)]
    def hcol(pred):
        return next((k for k, h in enumerate(head) if pred(h)), None)
    c_nom, c_car, c_est = hcol(lambda h: h.startswith("NOMBRE")), hcol(lambda h: h == "CARGO"), hcol(lambda h: h == "ESTADO")
    c_ing = hcol(lambda h: h.startswith("FECHA DE INGRESO"))
    c_fin = hcol(lambda h: h.startswith("FECHA DE FIN"))
    c_pto = hcol(lambda h: h == "PUNTO")
    for r in rows_from(ws_ap, 2):
        nombre = g(r, c_nom) if c_nom is not None else None
        if not nombre or not str(nombre).strip():
            continue
        m = match_nomina(nombre)
        out["aprendices"].append({
            "nombre": str(nombre).strip(),
            "programa": str(g(r, c_car)).strip() if c_car is not None and g(r, c_car) else None,
            "estado": str(g(r, c_est)).strip().upper() if c_est is not None and g(r, c_est) else None,
            "fechaIngreso": as_date(g(r, c_ing)) if c_ing is not None else None,
            "fechaFin": as_date(g(r, c_fin)) if c_fin is not None else None,
            "sede": (g(r, c_pto) if c_pto is not None else None) or (m["co"] if m else None),
            "area": m["centroCosto"] if m else None,
            "cargoNomina": m["cargo"] if m else None,
        })
out["aprendices"].sort(key=lambda x: x["fechaFin"] or "9999")

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


wb_rot = openpyxl.load_workbook(os.path.join(DATA_DIR, buscar_xlsx("rotacion")), data_only=True)
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
# ordene, filtre y cruce bien con el detalle de CAPACITACIONES. También "julio" -> "Julio".
for _lista in out.values():
    for _r in _lista:
        if isinstance(_r.get("mes"), str):
            _r["mes"] = _r["mes"].strip()
            if _r["mes"].capitalize() in MESES:
                _r["mes"] = _r["mes"].capitalize()
out["formacion"].sort(key=lambda x: period_key(x["anio"], x["mes"]))
out["costo"].sort(key=lambda x: period_key(x["anio"], x["mes"]))

out_path = os.path.join(BASE, "gestion_humana_data.json")
with open(out_path, "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, indent=2)

print("OK ->", out_path)
for k, v in out.items():
    print(f"  {k}: {len(v)} filas")
