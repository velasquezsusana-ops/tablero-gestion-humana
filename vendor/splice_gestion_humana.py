# -*- coding: utf-8 -*-
"""
Hornea gestion_humana_data.json dentro de "Tablero Gestion Humana.html":
reemplaza el bloque `const DATA_RAW = {...};` (y las líneas de reviveDate()
que lo acompañan) por los datos frescos extraídos de los 4 excels fuente.

Localiza el bloque por marcadores de texto (no por número de línea), así que
es seguro volver a correrlo aunque el resto del HTML se haya editado.

Volver a correr esto (después de extract_gestion_humana.py) cada vez que se
actualicen los excels fuente.
"""
import json
import os

BASE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.dirname(BASE)
HTML_PATH = os.path.join(DATA_DIR, "Tablero Gestion Humana.html")
JSON_PATH = os.path.join(BASE, "gestion_humana_data.json")

with open(HTML_PATH, "r", encoding="utf-8") as f:
    lines = f.readlines()

with open(JSON_PATH, "r", encoding="utf-8") as f:
    data = json.load(f)

data_compact = json.dumps(data, ensure_ascii=False, separators=(",", ":"))

START_MARKER = " *  Dataset — horneado desde"
END_MARKER = "let DATA = DATA_RAW;"

start_idx = next(i for i, l in enumerate(lines) if l.strip().startswith("/* " + "-"*10) and START_MARKER in lines[i+1])
end_idx = next(i for i, l in enumerate(lines) if END_MARKER in l)

boot_block = """/* ---------------------------------------------------------------- *
 *  Dataset — horneado desde "Gestión humana.xlsx", "Informe tipo de
 *  contrato.xlsx", "INGRESOS.xlsx" y "ROTACION..xlsx" por
 *  extract_gestion_humana.py + splice_gestion_humana.py. No se lee en vivo:
 *  para actualizar, correr esos dos scripts de nuevo.
 * ---------------------------------------------------------------- */
function reviveDate(iso){
  if (!iso) return null;
  const [y,m,d] = iso.split('-').map(Number);
  return new Date(y, m-1, d);
}
const DATA_RAW = %s;
DATA_RAW.nomina.forEach(r=>{ r.fechaIngreso = reviveDate(r.fechaIngreso); r.fechaContratoHasta = reviveDate(r.fechaContratoHasta); });
DATA_RAW.ingresos.forEach(r=>{ r.fechaIngreso = reviveDate(r.fechaIngreso); r.fechaSolicitud = reviveDate(r.fechaSolicitud); r.fechaMaxima = reviveDate(r.fechaMaxima); });
DATA_RAW.sst.forEach(r=>{ r.fecha = reviveDate(r.fecha); });
let DATA = DATA_RAW;
""" % (data_compact,)

out = lines[0:start_idx] + [boot_block] + lines[end_idx+1:]

with open(HTML_PATH, "w", encoding="utf-8") as f:
    f.writelines(out)

print("OK -> spliced", HTML_PATH)
print("HTML size:", os.path.getsize(HTML_PATH), "bytes")
