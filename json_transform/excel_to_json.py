import json
import re
from pathlib import Path

import openpyxl
import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / 'input'
OUTPUT_DIR = BASE_DIR / 'output'

ruta_excel = INPUT_DIR / 'data.xlsx'

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def nombre_archivo_seguro(nombre_hoja: str) -> str:
    """Convierte el nombre de la hoja en un nombre de archivo válido."""
    nombre = nombre_hoja.strip() or 'hoja'
    nombre = re.sub(r'[<>:"/\\|?*]', '_', nombre)
    return f'{nombre}.json'


wb = openpyxl.load_workbook(ruta_excel, read_only=True)
hojas_visibles = [ws.title for ws in wb.worksheets if ws.sheet_state == 'visible']
wb.close()

hojas = pd.read_excel(
    ruta_excel,
    sheet_name=hojas_visibles or None,
    engine='openpyxl',
)
# Si solo hay una hoja, pandas devuelve DataFrame; unificar a dict
if isinstance(hojas, pd.DataFrame):
    hojas = {hojas_visibles[0]: hojas}

for nombre_hoja, df in hojas.items():
    data = json.loads(df.to_json(orient='records', force_ascii=False, date_format='iso'))
    ruta_json = OUTPUT_DIR / nombre_archivo_seguro(nombre_hoja)

    with open(ruta_json, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

    print(f'Archivo JSON generado: {ruta_json}')

print(f'Total de hojas convertidas: {len(hojas)}')
