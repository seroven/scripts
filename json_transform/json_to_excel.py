import json
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / 'input'
OUTPUT_DIR = BASE_DIR / 'output'

ruta_json = INPUT_DIR / 'data.json'
ruta_excel = OUTPUT_DIR / 'data.xlsx'

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

with open(ruta_json, 'r', encoding='utf-8') as f:
    data = json.load(f)

df = pd.DataFrame(data)
df.to_excel(ruta_excel, index=False, engine='openpyxl')

print(f'Archivo Excel generado: {ruta_excel}')
