"""
prepare_inputs_for_calc_v9.py
================================
STAGE 1 of the "run v9, then align to 34" pipeline.

Run this BEFORE calc_engine_v9.py. It takes the real module-34
Inputs_v2.csv + cso_rates.csv and writes, into the same directory:

  - Inputs_v2.csv        overwritten with a v9-COMPATIBLE version
  - cso_rates.csv         passed through unchanged (already compatible)
  - mortality_v2.csv      dummy placeholder (see WHY DUMMY FILES below)
  - cv_rates_v2.csv       dummy placeholder
  - term_rates.csv        dummy placeholder
  - dividend_assumptions.csv   dummy placeholder

so that calc_engine_v9.py runs to completion without a Python-level
crash (missing file / KeyError), on module-34 data it was never
written for.

WHY v9's ROW ORDER MUST BE PRESERVED (read this before editing)
---------------------------------------------------------------------
v9 mostly looks up Inputs values by LABEL (via its ifinp()/inp()
helpers, which do INDEX/MATCH on column A text) -- that part is
robust to row order. But three of its formula blocks (CV_RATES's
"FULL EXACT STRUCTURE" section) hard-code cell positions instead:

    Inputs!$B$8   -- v9 means "Paid-Up Age"
    Inputs!$B$19  -- v9 means "Expense Loading Factor"
    Inputs!$B$21  -- v9 means "Accumulation Interest Rate (p.a.)"

Module 34's own natural row layout puts different things at those
exact rows (row 8 = "Premium Paying Term (PPT)", row 19 = "Dividend
Option", row 21 = "Annual Base Premium"). If we wrote module 34's data
using module 34's own row order, v9 would silently feed e.g. a PPT
value of "30" into a formula that thinks it's reading an age -- wrong
numbers, no error raised.

So this script writes Inputs_v2.csv using v9's OWN 27-row layout
(verified against your uploaded WL_model_31_32.xlsx, which is a real
v9 output), placing every module-34 value at the row v9 expects for
that label, and inserting two rows that don't exist in module 34 at
all -- "Paid-Up Age" and "Expense Loading Factor" -- with dummy values
chosen to be actuarially inert rather than merely "a number":

    Paid-Up Age = 121 (max mortality table age): any "IF(age >=
        paid_up_age, ...)" branch in v9's approximated formulas never
        triggers -- which is the CORRECT behavior for a PPT-based
        product that has no paid-up-age concept at all, not just a
        convenient placeholder.
    Expense Loading Factor = 0: neutral, adds no expense loading.

IMPORTANT -- WHAT THIS DOES **NOT** GUARANTEE
--------------------------------------------------
The dummy values above make v9 run without crashing. They do NOT make
v9's CV_RATES / Commutation / Projection_Monthly formulas actuarially
correct for module 34 -- v9's formulas encode 31/32's paid-up-at-age
product logic, which is a different product design than module 34's
PPT-based logic. That's expected and fine here, because
align_to_module_34.py (stage 3) completely overwrites every
calculation sheet's formulas afterward with formulas verified against
your real module 34 reference workbook. Nothing calculated by v9 in
this intermediate file survives into the final output -- v9 is being
used purely as a workbook-shell generator in this pipeline, not as
the source of any final numbers.

WHY DUMMY DATA-TABLE CSVs (mortality_v2 / cv_rates_v2 / term_rates /
dividend_assumptions)
--------------------------------------------------------------------
calc_engine_v9.py's CSV_INPUTS dict names these four files and reads
them with write_csv_to_sheet() before doing anything else -- if they
don't exist on disk, that's a Python FileNotFoundError before v9 gets
anywhere near your actual data. Their CONTENT doesn't matter:
  - CV_RATES and TERM_RATES get fully overwritten later in v9's own
    script by its "FULL EXACT STRUCTURE" formula-writing loops -- the
    CSV-loaded content is discarded within v9 itself.
  - MORTALITY and DIvidend_Assumption keep whatever the CSV had, but
    align_to_module_34.py overwrites both anyway.
So a tiny placeholder grid is enough; this script writes a small one.

APPLY_REFERENCE_SHEET_FIXES CAVEAT
--------------------------------------
calc_engine_v9.py also calls apply_reference_sheet_fixes(wb), which
loads a file named WL_Model_Paidup_31_32_Final.xlsx to clone exact
ETI/APL/RPU/Loan sheets from. That file is specific to your 31/32
setup; nothing here can produce or substitute for it. If it isn't
present when you run v9 for a module-34 case, v9 will raise
FileNotFoundError at that line. Since align_to_module_34.py rebuilds
those same four sheets from verified module-34 data anyway, the
simplest fix is to skip that call for this pipeline -- comment out (or
guard with a try/except FileNotFoundError: pass) the line:
    apply_reference_sheet_fixes(wb)
near the end of calc_engine_v9.py. Nothing downstream depends on it.

USAGE
-----
    python prepare_inputs_for_calc_v9.py \\
        --inputs Inputs_v2.csv \\
        --cso cso_rates.csv \\
        --outdir .

Then run calc_engine_v9.py as usual (it reads the filenames this
script just wrote, from the same directory).
"""

import argparse
import csv
import os

# ----------------------------------------------------------------------
# v9's expected 27-row 'Inputs' layout (label -> row), reverse-engineered
# from calc_engine_v9.py's hardcoded Inputs!$B$N references and confirmed
# against the actual row positions in your uploaded WL_model_31_32.xlsx.
# ----------------------------------------------------------------------
V9_ROW_ORDER = [
    "Mortality Table", "Underwriting Class", "Plan Code", "Policy Number",
    "Gender", "Issue Age", "Paid-Up Age", "Units", "PUA Units", "Issue Date",
    "Valuation Date", "DOB", "Current age", "Duration(x)", "Duration(x+1)",
    "Base SA", "Declared Dividend Rate (p.a.)", "Expense Loading Factor",
    "Dividend Option", "Accumulation Interest Rate (p.a.)", "Annual Base Premium",
    "Modal Premium", "Premium Mode", "Modal Factor", "Contract Status",
    "Premium Paid-Up Duration", "Is Currently Paying Premium",
]

# Labels that exist in module 34's Inputs_v2.csv but not in v9's layout at
# all (v9 never looks them up by label either -- safe to just drop).
MODULE_34_ONLY_LABELS = {"Premium Paying Term (PPT)"}

# Labels v9 expects that module 34 has no concept of -- filled with dummy
# values chosen to be actuarially inert (see module docstring).
DUMMY_VALUES = {
    "Paid-Up Age": "121",
    "Expense Loading Factor": "0",
}

# Formula cells (position-dependent on v9's OWN row numbers, all of which
# happen to match module 34's row numbers for these particular labels --
# see the module docstring's row-collision table for why only these four
# are safe to write as live formulas).
FORMULA_CELLS = {
    "Issue Age": "=ROUND((B11-B13)/365.25,0)",
    "Current age": "=INT((B12-B13)/365.25)",
    "Duration(x)": "=INT((B12-B11)/365.25)",
    "Duration(x+1)": "=INT((B12-B11)/365.25)+1",
}


EXCEL_ERROR_STRINGS = {"#NAME?", "#VALUE!", "#REF!", "#DIV/0!", "#N/A", "#NULL!", "#NUM!"}


def _clean_value(raw: str) -> str:
    """Normalize a raw CSV value before handing it to v9's own smart_cast():
    - '0.00%' style percent strings become plain decimal numbers, since
      v9's smart_cast has no percent handling and would otherwise store
      the '%' string as literal text instead of a number.
    - Leftover Excel error placeholders from the source export (e.g.
      '#NAME?' on Modal Premium/Modal Factor, which v9 never reads by
      label anyway) become '0' instead of being copied through as text.
    """
    raw = raw.strip()
    if raw in EXCEL_ERROR_STRINGS:
        return "0"
    if raw.endswith("%"):
        try:
            return str(float(raw[:-1].strip()) / 100.0)
        except ValueError:
            return raw
    return raw


def load_label_value_csv(path: str) -> dict:
    values = {}
    with open(path, newline="", encoding="utf-8-sig") as f:
        for row in csv.reader(f):
            if len(row) < 2:
                continue
            label, value = row[0].strip(), (row[1] or "").strip()
            if label and value != "":
                values[label] = _clean_value(value)
    return values


def build_v9_inputs_rows(module34_values: dict) -> list:
    rows = []
    for label in V9_ROW_ORDER:
        if label in FORMULA_CELLS:
            value = FORMULA_CELLS[label]
        elif label in DUMMY_VALUES:
            value = DUMMY_VALUES[label]
        elif label in module34_values:
            value = module34_values[label]
        else:
            value = "0"  # label exists in v9's layout but wasn't supplied -- inert default
        rows.append([label, value])

    unused = [l for l in module34_values if l in MODULE_34_ONLY_LABELS]
    if unused:
        print(f"[info] Dropped module-34-only labels not used by v9: {unused}")
    unmapped = [l for l in module34_values if l not in V9_ROW_ORDER and l not in MODULE_34_ONLY_LABELS]
    if unmapped:
        print(f"[warn] Labels present in your Inputs_v2.csv but not recognized "
              f"by either v9 or module 34's layout -- ignored: {unmapped}")
    return rows


def write_label_value_csv(path: str, rows: list):
    """No header row -- v9's write_csv_to_sheet() writes the Inputs sheet
    starting at Excel row 2 directly from the CSV's first line, so a
    header row here would shift every single row down by one and break
    every hardcoded and label-based lookup in v9."""
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerows(rows)


def write_dummy_grid(path: str, rows: int = 5, cols: int = 5):
    if os.path.exists(path):
        return  # don't clobber a real file if one already happens to exist
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        for _ in range(rows):
            w.writerow([0] * cols)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--inputs", required=True, help="Path to module 34's Inputs_v2.csv (label,value pairs)")
    ap.add_argument("--cso", required=True, help="Path to module 34's cso_rates.csv")
    ap.add_argument("--outdir", default=".", help="Directory calc_engine_v9.py will read from (default: cwd)")
    args = ap.parse_args()

    os.makedirs(args.outdir, exist_ok=True)

    module34_values = load_label_value_csv(args.inputs)
    v9_rows = build_v9_inputs_rows(module34_values)
    out_inputs_path = os.path.join(args.outdir, "Inputs_v2.csv")
    write_label_value_csv(out_inputs_path, v9_rows)
    print(f"[1/6] Wrote v9-compatible Inputs_v2.csv -> {out_inputs_path} ({len(v9_rows)} rows, no header)")

    # cso_rates.csv passes through unchanged -- same format v9 already expects.
    out_cso_path = os.path.join(args.outdir, "cso_rates.csv")
    if os.path.abspath(args.cso) != os.path.abspath(out_cso_path):
        with open(args.cso, "rb") as src, open(out_cso_path, "wb") as dst:
            dst.write(src.read())
    print(f"[2/6] cso_rates.csv ready -> {out_cso_path}")

    dummies = ["mortality_v2.csv", "cv_rates_v2.csv", "term_rates.csv", "dividend_assumptions.csv"]
    for i, name in enumerate(dummies, start=3):
        p = os.path.join(args.outdir, name)
        write_dummy_grid(p)
        print(f"[{i}/6] Dummy placeholder ready -> {p} (content is discarded/overwritten later, see docstring)")

    print("\nDone. calc_engine_v9.py can now be run from this directory.")
    print("Reminder: if WL_Model_Paidup_31_32_Final.xlsx isn't present, v9 will")
    print("stop at apply_reference_sheet_fixes(wb) -- see module docstring for the one-line fix.")


if __name__ == "__main__":
    main()
