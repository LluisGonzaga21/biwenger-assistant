"""
Generates the analysis Excel file: market + rival squads + your squad,
with points/market-value and points/clause ratios (current season and,
if the API provides it, the previous season).

IMPORTANT: the Biwenger API is unofficial and undocumented, so the field
names (see CAMPOS_CANDIDATOS in biwenger_helpers.py) are the best
approximation possible based on community clients. If you see "field not
found" warnings when running it, run 01_explorar_api.py first, look at
the real JSON in output/debug/, and we'll adjust those lists together.

Uses the same local cache as the analysis notebook (output/cache/*.json)
-- the first time it downloads everything from the API (it can take
several minutes due to the request rate limit), subsequent runs are
practically instant.

Usage:
    1. Copy .env.example to .env and fill it in.
    2. (recommended the first time) python 01_explorar_api.py  -> validates the connection
    3. python 02_generar_excel.py            -> uses the cache if it already exists
       python 02_generar_excel.py --refresh  -> re-downloads everything from the API
    4. The Excel file is generated at output/analisis_biwenger_<date>.xlsx
"""

from __future__ import annotations

import os
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter

from biwenger_helpers import load_league_data, build_dataframe

# prevents a non-ASCII character in a print from crashing the script on
# consoles that don't use UTF-8 (e.g. cmd.exe with the default codepage) --
# it replaces the character instead of raising UnicodeEncodeError
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(errors="replace")

load_dotenv()

EMAIL = os.getenv("BIWENGER_EMAIL")
PASSWORD = os.getenv("BIWENGER_PASSWORD")
LEAGUE_ID = os.getenv("BIWENGER_LEAGUE_ID") or None
OWN_TEAM_ID = os.getenv("BIWENGER_OWN_TEAM_ID") or None  # manual override if auto-detect fails

OUTPUT_DIR = Path(__file__).parent / "output"

COLUMNAS_RATIO = [
    "Ratio pts/VM (actual)", "Ratio pts/VM (anterior)",
    "Ratio pts totales/clausula (actual)", "Ratio pts totales/clausula (anterior)",
    "Ratio pts medios/clausula (actual)", "Ratio pts medios/clausula (anterior)",
]


def format_sheet(writer, df: pd.DataFrame, nombre_hoja: str, columnas_ratio: list):
    if df.empty:
        df = pd.DataFrame({"Sin datos": []})
    elif "Jugador" in df.columns:
        # make 'Jugador' the first column, so it can be frozen and the
        # name stays visible even when scrolling right
        resto = [c for c in df.columns if c != "Jugador"]
        df = df[["Jugador"] + resto]

    df.to_excel(writer, sheet_name=nombre_hoja, index=False)
    ws = writer.sheets[nombre_hoja]

    # freeze the header row and the 'Jugador' column (column A) at the same time
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, col_name in enumerate(df.columns, start=1):
        max_len = max([len(str(col_name))] + [len(str(v)) for v in df[col_name].astype(str).tolist()[:200]])
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 10), 40)

    for col_name in columnas_ratio:
        if col_name not in df.columns:
            continue
        col_idx = df.columns.get_loc(col_name) + 1
        col_letter = get_column_letter(col_idx)
        rango = f"{col_letter}2:{col_letter}{len(df) + 1}"
        rule = ColorScaleRule(
            start_type="min", start_color="F8696B",
            mid_type="percentile", mid_value=50, mid_color="FFEB84",
            end_type="max", end_color="63BE7B",
        )
        ws.conditional_formatting.add(rango, rule)


def main():
    if not EMAIL or not PASSWORD:
        sys.exit(
            "Missing BIWENGER_EMAIL / BIWENGER_PASSWORD.\n"
            "Copy .env.example to .env and fill it in with your Biwenger credentials."
        )

    forzar_refresco = "--refresh" in sys.argv
    try:
        data = load_league_data(
            EMAIL, PASSWORD, league_id=LEAGUE_ID, own_team_id=OWN_TEAM_ID,
            forzar_refresco=forzar_refresco,
        )
    except RuntimeError as e:
        sys.exit(str(e))

    print(f"league: {LEAGUE_ID or '(auto)'}, your team_id: {data['mi_team_id']}, "
          f"score_id: {data['score_id']}, balance: {data['balance']:,}")

    df = build_dataframe(data)

    df_mercado = df[df["Origen"] == "Mercado"]
    df_rivales = df[df["Origen"].str.startswith("Rival:")]
    df_mias = df[df["Origen"] == "Mi equipo"]

    if not df_mercado.empty and "Ratio pts/VM (actual)" in df_mercado:
        df_mercado = df_mercado.sort_values("Ratio pts/VM (actual)", ascending=False)
    if not df_rivales.empty and "Ratio pts/VM (actual)" in df_rivales:
        df_rivales = df_rivales.sort_values("Ratio pts/VM (actual)", ascending=False)
    if not df_mias.empty and "Ratio pts/VM (actual)" in df_mias:
        df_mias = df_mias.sort_values("Ratio pts/VM (actual)", ascending=False)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    fecha = datetime.now().strftime("%Y%m%d_%H%M")
    out_path = OUTPUT_DIR / f"analisis_biwenger_{fecha}.xlsx"

    print(f"Generating Excel at {out_path}...")
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        format_sheet(writer, df_mercado, "Mercado", COLUMNAS_RATIO)
        format_sheet(writer, df_rivales, "Rivales", COLUMNAS_RATIO)
        format_sheet(writer, df_mias, "Mi equipo", COLUMNAS_RATIO)

    print(f"\nDone -> {out_path}")


if __name__ == "__main__":
    main()
