import argparse
import pandas as pd
import sys
from pathlib import Path
import re
import json


def find_header_row(df):
    """Find a row index where a header like 'Betrag' appears."""
    for idx in range(min(10, len(df))):
        row = df.iloc[idx].astype(str).str.lower().fillna("")
        if any('betrag' in str(x) for x in row.values):
            return idx
    return None


def extract_sheet_transactions(df_sheet):
    """Given a sheet as a DataFrame (header=None), return extracted rows as list of dicts and metadata."""
    # normalize to string to ease searching
    df = df_sheet.copy()

    # find header row where 'Betrag' exists
    hdr_idx = find_header_row(df)
    if hdr_idx is None:
        return []

    header = df.iloc[hdr_idx].fillna("").astype(str).str.strip().tolist()
    data = df.iloc[hdr_idx + 1 :].copy()
    data.columns = header

    # stop when encountering common footer markers
    stop_keywords = ['buchungsvermerke', 'kontodaten', 'verwendungszweck', 'erstellt', 'genehmigt']
    rows = []
    for _, r in data.iterrows():
        # check stop
        joined = ' '.join([str(x).lower() for x in r.values if pd.notna(x)])
        if any(k in joined for k in stop_keywords):
            break
        # consider row valid if any of key columns has value
        if r.isnull().all():
            continue
        rows.append(r.to_dict())

    return rows


def find_meta_values(df_sheet):
    """Search entire sheet for Datum: and Buchungsbeleg-Nr. patterns."""
    text = '\n'.join(df_sheet.fillna('').astype(str).stack().tolist())
    datum = None
    beleg = None

    m = re.search(r'Datum[:\s]+([0-9]{4}-[0-9]{2}-[0-9]{2}|[0-9]{1,2}\.[0-9]{1,2}\.[0-9]{2,4}|[0-9]{4}\/[0-9]{1,2}\/[0-9]{1,2}|[0-9]{1,2}\/[0-9]{1,2}\/[0-9]{2,4})', text, re.IGNORECASE)
    if m:
        datum = m.group(1)

    m2 = re.search(r'Buchungsbeleg-?Nr[:\.\s]*([0-9A-Za-z\-_/]+)', text, re.IGNORECASE)
    if m2:
        beleg = m2.group(1).strip()

    return datum, beleg


def load_column_mapping(config_path: Path = None):
    """Load column mapping from JSON config file if it exists."""
    if config_path is None:
        config_path = Path.cwd() / "column_mapping.json"
    if config_path.exists():
        try:
            with open(config_path) as f:
                return json.load(f)
        except Exception as e:
            print(f"Warning: could not load config {config_path}: {e}")
    return {}


def transform(template_path: Path, source_path: Path, output_path: Path, config_path: Path = None, defaults: dict = None):
    try:
        template_cols = list(pd.read_excel(template_path, nrows=0).columns)
    except Exception as e:
        print(f"Error reading template header: {e}")
        sys.exit(2)
    
    # Load column mapping config
    col_mapping = load_column_mapping(config_path)

    try:
        xls = pd.ExcelFile(source_path)
    except Exception as e:
        print(f"Error reading source file: {e}")
        sys.exit(3)

    out_rows = []
    for sheet in xls.sheet_names:
        try:
            sheet_df = pd.read_excel(xls, sheet_name=sheet, header=None, dtype=object)
        except Exception:
            continue

        datum, beleg = find_meta_values(sheet_df)
        extracted = extract_sheet_transactions(sheet_df)

        # map extracted rows into template columns
        for r in extracted:
            out = {c: pd.NA for c in template_cols}
            # Attempt common mappings
            # keys in r may include 'Betrag', 'Konto', 'Kreditor', 'Text', 'Beschreibung'
            keymap = {k.lower().strip(): k for k in r.keys()}
            def get_key(*candidates):
                for cand in candidates:
                    if cand.lower() in keymap:
                        return r[keymap[cand.lower()]]
                return pd.NA

            betrag_val = get_key('Betrag', 'betrag')
            if pd.notna(betrag_val):
                try:
                    betrag_val = get_key('Betrag', 'betrag')
                    if pd.notna(betrag_val):
                        try:
                            # convert to float, multiply by 100 and store as int
                            num = float(betrag_val)
                            out['BETRAG'] = int(round(num * 100))
                        except (ValueError, TypeError):
                            out['BETRAG'] = betrag_val
                except (ValueError, TypeError):
                    out['BETRAG'] = betrag_val
            
            konto_val = get_key('Konto', 'konto')
            if pd.notna(konto_val):
                # assume Konto maps to HABENKONTO
                if 'HABENKONTO' in out:
                    out['HABENKONTO'] = konto_val
                elif 'SOLLKONTO' in out:
                    out['SOLLKONTO'] = konto_val

            out['DEBI_KREDI'] = get_key('Kreditor', 'kreditor')
            # text may be under several headers or as the first column
            out['BUCH_TEXT'] = get_key('Text', 'Beschreibung', 'text')
            
            # KST -> KOSTSTELLE, KTR -> KOSTTRAGER
            if 'KOSTSTELLE' in out:
                out['KOSTSTELLE'] = get_key('KST', 'Koststelle', 'KOSTSTELLE')
            if 'KOSTTRAGER' in out:
                ktr = get_key('KTR', 'Kostträger', 'KOSTTRAGER')
                if pd.notna(ktr):
                    out['KOSTTRAGER'] = str(ktr).replace(' ', '')
                else:
                    out['KOSTTRAGER'] = ktr
            
            if pd.notna(datum) and 'BELEG_DAT' in out:
                out['BELEG_DAT'] = datum
            if pd.notna(beleg) and 'BELEG_NR' in out:
                out['BELEG_NR'] = beleg

            # fill defaults for SATZART, FIRMA, SOLL_HABEN, BUCH_KREIS, BUCH_JAHR, BUCH_MONAT
            if defaults:
                for key in ['SATZART', 'FIRMA', 'SOLL_HABEN', 'BUCH_KREIS', 'BUCH_JAHR', 'BUCH_MONAT']:
                    if key in out and key in defaults and defaults.get(key) is not None:
                        out[key] = defaults.get(key)

            # Format BUCH_TEXT: prefix = BUCH_MONAT (2 digits) + last 2 of BUCH_JAHR
            # and description is the substring starting with 'Bereitschaft'
            prefix = None
            try:
                if defaults and defaults.get('BUCH_MONAT') is not None and defaults.get('BUCH_JAHR') is not None:
                    mon = str(int(defaults.get('BUCH_MONAT'))).zfill(2)
                    yr = str(defaults.get('BUCH_JAHR'))
                    prefix = f"{mon}{yr[-2:]}"
            except Exception:
                prefix = None

            desc = out.get('BUCH_TEXT', '')
            s = str(desc)
            # Try to extract a substring that matches variants like "Bereitschaftspflege" (allow misspellings)
            mdesc = re.search(r'(Bere\w*pflege.*)', s, flags=re.IGNORECASE)
            if mdesc:
                s = mdesc.group(1).strip()
            else:
                # fallback: find any word starting with 'Bere' and take from there
                m2 = re.search(r'(Bere\w.*)', s, flags=re.IGNORECASE)
                if m2:
                    s = m2.group(1).strip()
                else:
                    s = s.strip()

            if prefix and s:
                out['BUCH_TEXT'] = f"{prefix} {s}"
            else:
                out['BUCH_TEXT'] = s

            out_rows.append(out)

    if not out_rows:
        print('No transactions extracted by heuristic.')

    output_df = pd.DataFrame(out_rows, columns=template_cols)

    # Post-process columns
    if 'BETRAG' in output_df.columns:
        output_df['BETRAG'] = pd.to_numeric(output_df['BETRAG'], errors='coerce')
        def to_int_val(v):
            try:
                if pd.isna(v):
                    return pd.NA
                return int(round(float(v)))
            except Exception:
                return pd.NA
        output_df['BETRAG'] = output_df['BETRAG'].apply(to_int_val).astype('Int64')

    if 'KOSTTRAGER' in output_df.columns:
        output_df['KOSTTRAGER'] = output_df['KOSTTRAGER'].fillna('').astype(str).str.replace(r'\s+', '', regex=True)
        output_df.loc[output_df['KOSTTRAGER'] == '', 'KOSTTRAGER'] = pd.NA

    if 'KOSTSTELLE' in output_df.columns:
        output_df['KOSTSTELLE'] = output_df['KOSTSTELLE'].fillna('').astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        output_df.loc[output_df['KOSTSTELLE'] == '', 'KOSTSTELLE'] = pd.NA

    # Format BELEG_DAT as YYYYMMDD strings when possible
    if 'BELEG_DAT' in output_df.columns:
        try:
            bd = pd.to_datetime(output_df['BELEG_DAT'], errors='coerce')
            output_df['BELEG_DAT'] = bd.dt.strftime('%Y%m%d')
            output_df.loc[output_df['BELEG_DAT'].isna(), 'BELEG_DAT'] = pd.NA
        except Exception:
            pass

    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        # remove existing file if possible (avoid permission error when file closed by us)
        try:
            if output_path.exists():
                output_path.unlink()
        except Exception:
            pass
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            output_df.to_excel(writer, index=False, sheet_name="Sheet1")
    except Exception as e:
        print(f"Error writing output file: {e}")
        sys.exit(4)

    print(f"Created: {output_path.resolve()} with {len(output_df)} rows")

    # Non-interactive rename if provided via defaults['RENAME']
    if defaults and isinstance(defaults, dict) and defaults.get('RENAME'):
        try:
            new_name = defaults.get('RENAME')
            new_path = Path(new_name)
            if not new_path.suffix:
                new_path = output_path.parent / (new_name + output_path.suffix)
            elif not new_path.is_absolute():
                new_path = output_path.parent / new_name
            if new_path.exists():
                try:
                    new_path.unlink()
                except Exception:
                    pass
            output_path.replace(new_path)
            output_path = new_path
            print(f'Renamed output to: {output_path.resolve()}')
        except Exception as e:
            print(f'Could not rename file (non-interactive): {e}')

    # interactive rename: ask user if they want to rename the output file
    try:
        # if defaults contains a key to skip rename, respect it (defaults may be from CLI prompting)
        skip_rename = False
    except Exception:
        skip_rename = False

    # The CLI top-level will add a --no-rename flag and pass through via args; however
    # when transform() is called programmatically, the caller may pass defaults with 'NO_RENAME'.
    if defaults and isinstance(defaults, dict) and defaults.get('NO_RENAME'):
        skip_rename = True

    if not skip_rename:
        try:
            new_name = input('Enter new output filename (or press Enter to keep current): ').strip()
            if new_name:
                new_path = Path(new_name)
                if not new_path.suffix:
                    new_path = output_path.parent / (new_name + output_path.suffix)
                elif not new_path.is_absolute():
                    new_path = output_path.parent / new_name
                try:
                    # overwrite if exists
                    if new_path.exists():
                        new_path.unlink()
                    output_path.replace(new_path)
                    output_path = new_path
                    print(f'Renamed output to: {output_path.resolve()}')
                except Exception as e:
                    print(f'Could not rename file: {e}')
        except Exception:
            # non-interactive environment or input error
            pass


if __name__ == '__main__':
    p = argparse.ArgumentParser(description="Heuristic extractor: create a file similar to a template using data from a source Excel file.")
    p.add_argument("template", help="Path to template Excel file (e.g., 9241_1025_Bereitschatspflege_KRED.xlsx)")
    p.add_argument("source", help="Path to source Excel file (e.g., Auszahlungsbelege ...xlsx)")
    p.add_argument("--output", help="Output path (defaults to current folder with derived name)")
    p.add_argument("--config", help="Path to column_mapping.json (defaults to current folder)")
    p.add_argument("--satzart", help="SATZART value to fill")
    p.add_argument("--firma", help="FIRMA value to fill")
    p.add_argument("--soll_haben", help="SOLL_HABEN value to fill")
    p.add_argument("--buch_kreis", help="BUCH_KREIS value to fill")
    p.add_argument("--buch_jahr", help="BUCH_JAHR value to fill")
    p.add_argument("--buch_monat", help="BUCH_MONAT value to fill")
    p.add_argument("--no-rename", action='store_true', help="Do not prompt to rename output file after creation")
    p.add_argument("--rename", help="Non-interactive rename: provide the new output filename")
    args = p.parse_args()

    template_path = Path(args.template)
    source_path = Path(args.source)
    if args.output:
        output_path = Path(args.output)
    else:
        out_name = template_path.stem + "_from_" + source_path.stem + ".xlsx"
        output_path = Path.cwd() / out_name

    config_path = Path(args.config) if args.config else None
    # prepare defaults: use CLI args or prompt interactively
    def get_or_prompt(attr, prompt_label):
        val = getattr(args, attr)
        if val is None:
            try:
                v = input(f"{prompt_label} (press Enter to leave blank): ")
                val = v if v != "" else None
            except Exception:
                val = None
        return val

    defaults = {
        'SATZART': get_or_prompt('satzart', 'SATZART'),
        'FIRMA': get_or_prompt('firma', 'FIRMA'),
        'SOLL_HABEN': get_or_prompt('soll_haben', 'SOLL_HABEN'),
        'BUCH_KREIS': get_or_prompt('buch_kreis', 'BUCH_KREIS'),
        'BUCH_JAHR': get_or_prompt('buch_jahr', 'BUCH_JAHR'),
        'BUCH_MONAT': get_or_prompt('buch_monat', 'BUCH_MONAT'),
    }
    if args.no_rename:
        defaults['NO_RENAME'] = True
    if args.rename:
        defaults['RENAME'] = args.rename

    transform(template_path, source_path, output_path, config_path, defaults)
