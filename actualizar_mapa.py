import json
import os
import urllib.request
from collections import defaultdict
from datetime import date, datetime

# Enlace de solo lectura del libro de OneDrive.
SOURCE_URL = "https://1drv.ms/x/c/c244dfd4ed38b60a/IQBjETmP-kitSoMnUIXwu5lQAVcgyyzaq-vtg1_rXej9mqA?download=1"
FILE = "fuente.xlsx"

req = urllib.request.Request(SOURCE_URL, headers={"User-Agent": "Mozilla/5.0"})
with urllib.request.urlopen(req, timeout=60) as response:
    data = response.read()
    final_url = response.geturl()

# Si OneDrive entrega una página HTML en lugar del archivo, detener el proceso.
if not data[:4] == b"PK\x03\x04":
    raise RuntimeError("OneDrive no entregó el Excel directamente. URL final: " + final_url)

with open(FILE, "wb") as f:
    f.write(data)

from openpyxl import load_workbook
wb = load_workbook(FILE, read_only=True, data_only=True)
if "MapaPublico" not in wb.sheetnames:
    raise RuntimeError("No existe la hoja MapaPublico")
ws = wb["MapaPublico"]

headers = [str(c.value or "").strip().lower() for c in next(ws.iter_rows(min_row=1, max_row=1))]
aliases = {
    "codigo": ["codigo st", "código st", "código de sitio", "codigo de sitio"],
    "nombre": ["nombre de sitios", "nombre del sitio", "nombre de sitio"],
    "tarea": ["tarea"],
    "estado": ["estado"],
    "fecha": ["fecha de plan", "fecha plan"]
}
idx = {}
for key, names in aliases.items():
    found = next((headers.index(n) for n in names if n in headers), None)
    if found is None:
        raise RuntimeError("No se encontró la columna requerida: " + key + ". Encabezados: " + str(headers))
    idx[key] = found

groups = defaultdict(lambda: {"c":"", "n":"", "tareas":[], "estados":[], "fechas":[]})
for row in ws.iter_rows(min_row=2, values_only=True):
    code = str(row[idx["codigo"]] or "").strip().upper()
    if not code:
        continue
    item = groups[code]
    item["c"] = code
    item["n"] = str(row[idx["nombre"]] or "").strip()
    task = str(row[idx["tarea"]] or "").strip()
    state = str(row[idx["estado"]] or "").strip()
    planned = row[idx["fecha"]]
    if task and task not in item["tareas"]:
        item["tareas"].append(task)
    if state and state not in item["estados"]:
        item["estados"].append(state)
    if planned:
        if isinstance(planned, (datetime, date)):
            planned = planned.strftime("%d/%m/%Y")
        else:
            planned = str(planned).strip()
        if planned and planned not in item["fechas"]:
            item["fechas"].append(planned)

with open("datos_mapa.json", "w", encoding="utf-8") as f:
    json.dump(list(groups.values()), f, ensure_ascii=False, separators=(",", ":"))
print(f"Actualizados {len(groups)} sitios desde MapaPublico")
