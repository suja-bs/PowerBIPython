"""

pbix_to_dataframes.py

---------------------

Extract all tables from a Power BI .pbix file into a dict of pandas DataFrames.
 
Supports two data paths:

  1. DataMashup  — embedded Excel/CSV source files (Power Query imports)

  2. DataModel   — VertiPaq columnar store (the binary engine, read via pbixray)
 
Install:

    pip install pandas openpyxl xlrd pbixray

"""
 
import zipfile

import io

import os

from pathlib import Path

from typing import Dict
 
import pandas as pd
 
 
 
 
 
 
# ---------------------------------------------------------------------------

# DataMashup path (Excel / CSV imports)

# ---------------------------------------------------------------------------
 
def _read_excel_all_sheets(raw_bytes: bytes, source_name: str) -> Dict[str, pd.DataFrame]:

    buf = io.BytesIO(raw_bytes)

    try:

        return dict(pd.read_excel(buf, sheet_name=None, engine="openpyxl"))

    except Exception:

        buf.seek(0)

        try:

            return dict(pd.read_excel(buf, sheet_name=None, engine="xlrd"))

        except Exception as e:

            print(f"  [warn] Could not read {source_name}: {e}")

            return {}
 
 
def _read_csv_bytes(raw_bytes: bytes, source_name: str) -> pd.DataFrame | None:

    try:

        return pd.read_csv(io.BytesIO(raw_bytes))

    except Exception as e:

        print(f"  [warn] Could not read CSV {source_name}: {e}")

        return None
 
 
def _extract_from_datamashup(mashup_bytes: bytes) -> Dict[str, pd.DataFrame]:

    result: Dict[str, pd.DataFrame] = {}

    try:

        mashup_zip = zipfile.ZipFile(io.BytesIO(mashup_bytes))

    except zipfile.BadZipFile:

        print("  [warn] DataMashup is not a valid ZIP — skipping.")

        return result
 
    entries = mashup_zip.namelist()

    print(f"  DataMashup contains {len(entries)} entries.")
 
    for entry in entries:

        lower = entry.lower()

        if lower.endswith((".xlsx", ".xlsm", ".xls")):

            raw = mashup_zip.read(entry)

            sheets = _read_excel_all_sheets(raw, entry)

            for sheet_name, df in sheets.items():

                key = f"{Path(entry).stem}__{sheet_name}"

                result[key] = df

                print(f"    ✓ Excel sheet  →  '{key}'  shape={df.shape}")

        elif lower.endswith(".csv"):

            raw = mashup_zip.read(entry)

            df = _read_csv_bytes(raw, entry)

            if df is not None:

                key = Path(entry).stem

                result[key] = df

                print(f"    ✓ CSV          →  '{key}'  shape={df.shape}")
 
    mashup_zip.close()

    return result
 
 
# ---------------------------------------------------------------------------

# DataModel path (VertiPaq binary — via pbixray)

# ---------------------------------------------------------------------------
 
def _extract_from_datamodel(pbix_path: Path) -> Dict[str, pd.DataFrame]:

    try:

        from pbixray import PBIXRay

    except ImportError:

        print(

            "  [warn] pbixray not installed. Run: pip install pbixray\n"

            "         Skipping DataModel extraction."

        )

        return {}
 
    result: Dict[str, pd.DataFrame] = {}
 
    print("  Reading DataModel via pbixray (VertiPaq decoder)…")

    try:

        with PBIXRay(str(pbix_path)) as model:

            tables = model.tables

            if tables is None or len(tables) == 0:

                print("  [warn] No tables found in DataModel.")

                return result
 
            # Filter out Power BI internal/hidden tables (names starting with '$' or 'DateTableTemplate')

            all_table_names = list(tables)

            user_tables = [

                t for t in all_table_names

                if not t.startswith("$") and not t.startswith("DateTableTemplate")

            ]

            hidden = len(all_table_names) - len(user_tables)

            if hidden:

                print(f"  (skipping {hidden} internal PBI system table(s))")
 
            print(f"  Found {len(user_tables)} user table(s): {user_tables}")
 
            for table_name in user_tables:

                try:

                    df = model.get_table(table_name)

                    if df is not None and not df.empty:

                        result[table_name] = df

                        print(f"    ✓ Table  →  '{table_name}'  shape={df.shape}")

                    else:

                        print(f"    ~ Table  '{table_name}' is empty — skipped.")

                except Exception as e:

                    print(f"    [warn] Could not decode table '{table_name}': {e}")
 
    except Exception as e:

        print(f"  [error] Failed to open DataModel: {e}")
 
    return result
 
 
# ---------------------------------------------------------------------------

# Top-level extractor

# ---------------------------------------------------------------------------
 
def pbix_to_dataframes(

    pbix_path: str | os.PathLike,

    prefer: str = "auto",          # "auto" | "datamodel" | "datamashup"

) -> Dict[str, pd.DataFrame]:

    """

    Extract all tabular data from a .pbix file.
 
    Parameters

    ----------

    pbix_path : str or Path

        Path to the .pbix file.

    prefer : str

        "auto"        — try DataMashup first; fall back to DataModel if empty.

        "datamodel"   — only read the VertiPaq DataModel (pbixray).

        "datamashup"  — only read embedded Excel/CSV sources.
 
    Returns

    -------

    dict[str, pd.DataFrame]

        Keys are table/sheet names.

    """

    pbix_path = Path(pbix_path)

    if not pbix_path.exists():

        raise FileNotFoundError(f"File not found: {pbix_path}")

    if pbix_path.suffix.lower() != ".pbix":

        raise ValueError(f"Expected a .pbix file, got: {pbix_path.suffix}")
 
    print(f"\n{'='*60}")

    print(f"Opening: {pbix_path.name}")

    print(f"{'='*60}")
 
    all_frames: Dict[str, pd.DataFrame] = {}
 
    with zipfile.ZipFile(pbix_path, "r") as zf:

        contents = set(zf.namelist())

        has_mashup = "DataMashup" in contents

        has_model  = "DataModel"  in contents

        print(f"DataMashup present: {has_mashup} | DataModel present: {has_model}\n")
 
        # ── DataMashup ───────────────────────────────────────────────────────

        if prefer in ("auto", "datamashup") and has_mashup:

            print("── Extracting from DataMashup (embedded files)…")

            mashup_bytes = zf.read("DataMashup")

            frames = _extract_from_datamashup(mashup_bytes)

            all_frames.update(frames)
 
    # ── DataModel (outside the with block — pbixray re-opens the file) ────

    if has_model and (prefer == "datamodel" or (prefer == "auto" and not all_frames)):

        print("\n── Extracting from DataModel (VertiPaq binary)…")

        frames = _extract_from_datamodel(pbix_path)

        all_frames.update(frames)
 
    # ── Summary ──────────────────────────────────────────────────────────────

    print(f"\n{'='*60}")

    if all_frames:

        print(f"✓ Extracted {len(all_frames)} DataFrame(s):")

        for k, df in all_frames.items():

            print(f"  '{k}': {df.shape[0]:,} rows × {df.shape[1]} cols")

    else:

        print("✗ No DataFrames extracted.")

        print(

            "\nFallback options if pbixray also fails:\n"

            "  • DAX Studio  → export tables as CSV (free desktop tool)\n"

            "  • Power BI Desktop → right-click table → 'Export data'\n"

            "  • XMLA endpoint / semantic-link (Fabric notebooks)\n"

        )

    print(f"{'='*60}\n")
 
    return all_frames
 
 
# ---------------------------------------------------------------------------

# CLI

# ---------------------------------------------------------------------------
 
if __name__ == "__main__":

    import argparse
 
    parser = argparse.ArgumentParser(

        description="Extract all tables from a .pbix file into DataFrames / CSVs."

    )

    parser.add_argument("pbix_file", help="Path to the .pbix file")

    parser.add_argument(

        "--prefer",

        choices=["auto", "datamodel", "datamashup"],

        default="auto",

        help="Which data path to use (default: auto)",

    )

    parser.add_argument(

        "--output-dir",

        default=r"D:\Python_QA_automation\Deval_data_pro\extracted",

        metavar="DIR",

        help="Save each DataFrame as a CSV in this directory.",

    )

    parser.add_argument(

        "--show-head",

        type=int,

        default=0,

        metavar="N",

        help="Print the first N rows of each DataFrame.",

    )

    args = parser.parse_args()
 
    frames = pbix_to_dataframes(args.pbix_file, prefer=args.prefer)
 
    if args.show_head > 0:

        for name, df in frames.items():

            print(f"\n── {name} (first {args.show_head} rows) ──")

            print(df.head(args.show_head).to_string(index=False))
 
    if args.output_dir:

        out = Path(args.output_dir)

        out.mkdir(parents=True, exist_ok=True)

        for name, df in frames.items():

            safe = name.replace("/", "_").replace("\\", "_").replace(" ", "_")

            path = out / f"{safe}.csv"

            df.to_csv(path, index=False)

            print(f"Saved: {path}")
 