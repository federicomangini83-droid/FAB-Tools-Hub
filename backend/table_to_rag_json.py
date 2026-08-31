"""
table_to_rag_json.py
Converte una tabella Excel (.xlsx) o CSV in uno ZIP contenente
un file JSON per ogni riga, nel formato:
  {
    "id": "1",
    "table": "dati.xlsx",
    "document_content": "Colonna1: valore; Colonna2: valore; ..."
  }
"""

import io
import json
import sys
import zipfile
from pathlib import Path

import pandas as pd


INPUT_FILE  = r"C:\Users\federico.mangini\Downloads\CONVERSION_FACTOR.xlsx"   # percorso file input
OUTPUT_ZIP  = r"C:\Users\federico.mangini\Downloads\CONVERSION_FACTOR.zip" # percorso zip di output
SHEET       = None   # es. "Foglio1" — None = primo foglio (xlsx); ignorato per csv
NAME_JSON = "CONVERSION_FACTOR"



def read_table(input_path: str, sheet) -> pd.DataFrame:
    path = Path(input_path)
    if not path.exists():
        print(f"Errore: file '{input_path}' non trovato.", file=sys.stderr)
        sys.exit(1)

    suffix = path.suffix.lower()
    if suffix in (".xlsx", ".xlsm", ".xls"):
        df = pd.read_excel(path, sheet_name=sheet or 0, dtype=str)
    elif suffix == ".csv":
        df = pd.read_csv(path, dtype=str)
    else:
        print(f"Errore: formato '{suffix}' non supportato.", file=sys.stderr)
        sys.exit(1)

    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all").reset_index(drop=True)
    return df


def row_to_text(row: pd.Series) -> str:
    parts = []
    for col, val in row.items():
        if pd.notna(val) and str(val).strip() != "":
            parts.append(f"{col}: {val}")
    return "; ".join(parts)


def convert(input_path: str, output_zip: str, sheet=None) -> None:
    df = read_table(input_path, sheet)
    source_name = Path(input_path).name
    base_name   = Path(input_path).stem   # "dati" da "dati.xlsx"

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for i, (_, row) in enumerate(df.iterrows(), start=1):
            text = row_to_text(row)
            if not text.strip():
                continue

            chunk = {
                "id": str(i),
                "table": source_name,
                "document_content": text
            }

            first_col_name  = str(df.columns[0]).strip().replace(" ", "_")
            first_col_value = str(row.iloc[0]).strip().replace(" ", "_")
            filename = f"{NAME_JSON}_{i}.json"
            zf.writestr(filename, json.dumps(chunk, ensure_ascii=False, indent=2))

    out = Path(output_zip)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_bytes(zip_buffer.getvalue())

    print(f"✓ {i} file JSON scritti in '{output_zip}'")
    print(f"\nEsempio '{base_name}_1.json':")
    print(json.dumps({
        "id": "1",
        "table": source_name,
        "document_content": text,
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    convert(INPUT_FILE, OUTPUT_ZIP, SHEET)