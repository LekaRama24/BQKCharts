"""
bqk.py — one command: raw survey export → all report charts.

    python bqk.py data/Dataset_Total_Final.xlsx --codebook data/Codebook.xlsx
    python bqk.py data/Dataset_Total_Final.xlsx --codebook data/Codebook.xlsx --outliers 52 144 147 46 41

Result: folder grafiket_<year>/ with every chart as PNG and grafiket_<year>.zip with all of them.
The data is cleaned on the way (same rules as the report workbook), but the cleaned file is not kept.
"""
import argparse
import os
import shutil
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path[:0] = [os.path.join(HERE, "pastrimi"), os.path.join(HERE, "grafiket")]

import pastro_dataset  # noqa: E402
import figurat         # noqa: E402


def run(raw, dalja=None, codebook=None, viti=None, outliers=(), pragu=None, korrigjime=None, fleta=0,
        periudha="korrik–gusht", skema=None, pika=None):
    """Raw export → charts. The cleaned workbook is only an intermediate step (temporary folder, deleted).
    Returns (cleaning summary, chart folder, charts made, charts skipped, zip path)."""
    tmp = tempfile.mkdtemp()
    try:
        book = os.path.join(tmp, "dataset.xlsx")
        r = pastro_dataset.pastro(raw, dalja=book, viti=viti, pragu=pragu, outliers=outliers, korrigjime=korrigjime,
                                  codebook=codebook, skema=skema, fleta=fleta, pika=pika)
        year = r["viti"]
        out = dalja or f"grafiket_{year}"
        figdir, made, skipped = figurat.make_figures(book, dalja=out, viti=year, periudha=periudha)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    zpath = os.path.join(out, f"grafiket_{year}.zip")
    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
        for f in made:
            z.write(os.path.join(figdir, f), f)
    return r, out, made, skipped, zpath


def main():
    ap = argparse.ArgumentParser(description="Raw export → all report charts.")
    ap.add_argument("raw", help="raw survey export (.xlsx or .csv)")
    ap.add_argument("--codebook", help="codebook .xlsx (the sheet of the survey year is used)")
    ap.add_argument("--dalja", help="output folder (default: grafiket_<year>)")
    ap.add_argument("--viti", type=int, help="survey year (default: from the interview dates)")
    ap.add_argument("--outliers", nargs="*", type=int, default=[], help="'referues' numbers to mark as outliers")
    ap.add_argument("--pragu", type=float, help="threshold for 'shpenzime ≤ 10,000 €'")
    ap.add_argument("--korrigjime", help="CSV of manual corrections")
    ap.add_argument("--fleta", default=0, help="sheet of the raw export (default: the first)")
    ap.add_argument("--periudha", default="korrik–gusht", help="survey period for the chart source line")
    ap.add_argument("--skema", help="rules file (default: pastrimi/skema.json)")
    ap.add_argument("--pika", choices=["A", "M"], help="only if the export has no border-point column: A or M for all")
    a = ap.parse_args()
    a.fleta = int(a.fleta) if str(a.fleta).isdigit() else a.fleta

    r, out, made, skipped, zpath = run(**vars(a))
    print(f"\nViti {r['viti']}: {r['n_raw']} intervista → {r['n_db']} në DB ({r['n_A']} A, {r['n_M']} M), "
          f"{r['n_ex']} të përjashtuara")
    print(f"{len(made)} grafikë" + (f", {len(skipped)} të anashkaluar:" if skipped else ""))
    for f, why in skipped:
        print("   -", f, "→", why)
    if r["paralajmerime"]:
        print(f"\n{len(r['paralajmerime'])} paralajmërime për kontroll:")
        for w in r["paralajmerime"]:
            print("  ", w[1], "–", w[4])
    print(f"\nGati: {len(made)} grafikë në {out}/  (të gjithë në {zpath})")


if __name__ == "__main__":
    main()
