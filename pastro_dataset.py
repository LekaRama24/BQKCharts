"""
pastro_dataset.py — turns the raw survey export of any year into the final analysis workbook
("FINALEDataset_<year>_me_rezultatet_e_raportit.xlsx") with the sheets:

    Permbledhje   all report tables (live Excel formulas + stored results)
    Anketuesit    enumerator comparison
    DB            clean, valid interviews + derived columns + indicator block
    A, M          the same, per border point (Aeroport / Merdarë); a sheet per any other point (e.g. D in 2016)
    Perjashtuar   excluded interviews, with the reason
    Log           every change the script made (for review)

Usage (skema.json must be in the same folder as the script):
    python pastro_dataset.py Dataset_Total_Final.xlsx --codebook Codebook.xlsx
    python pastro_dataset.py Dataset_2018.xlsx --codebook Codebook.xlsx --viti 2018
    python pastro_dataset.py Dataset.xlsx --outliers 52 144 147 46 41 --pragu 10000 --korrigjime korrigjime.csv

Different years (all from skema.json + the codebook sheet of that year, nothing year-specific in the code):
  * column names: 'Pika', 'Date', 'Sa ditë', 'Akomodim', 'Total shudhëtimit', … are renamed to the 2026 names;
  * border point: letters (A/M), '1A'/'2M', or numbers whose meaning comes from the codebook (2018: 1 = Aeroport)
    or skema.json → always 1 = M, 2 = A as in 2026; other points (Dheu i Bardhë) get their own sheet;
  * Po/Jo as text, or Jo = 0 → 1/2;  remittance channels given in % → 1/0;
  * country of residence as ISO text (DE) → this year's numeric code; country/place labels from that year's codebook;
  * questions not asked in a year (pasaporta, depozita, blerje produktesh…) → empty column, table title marked KUJDES.

Cleaning rules (from the 2026 workbook):
  1. −999999999 ("nuk e di"/"nuk përgjigjet") → empty; counted per interview in 'Nr. përgjigjesh "nuk di" (Final)'.
  2. "##N/A##" (question skipped) → empty.
  3. Numbers typed as words in numeric fields ("asnje", "zero", "një", "pesë", …) → numbers; other text → empty (logged).
     Free-text fields (shpenzimeTjera, tetjera) are kept as typed.
  4. An interview date with fewer than 3 other interviews within ±5 days is treated as a typing error and replaced with
     the date of the neighbouring interviews (same assignment, or same enumerator) when they agree; else only flagged.
  5. Manual corrections from --korrigjime (CSV: interview__key,kolona,vlera,arsyeja) are applied last.
  6. Interviews with 0 or negative days in Kosovo are moved to Perjashtuar with the reason.

All formulas follow the 2026 workbook; ranges are sized to the number of rows of the year.
"""
import argparse
import csv
import json
import os
import re
import unicodedata
from datetime import datetime

import numpy as np
import pandas as pd
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

# =============================================================================================  settings
MISSING_CODES = [-999999999]
NA_STRINGS = ["##N/A##"]
TEXT_COLUMNS = ["interview__key", "interview__id", "shpenzimeTjera", "tetjera"]
DATE_COLUMNS = ["datapytesorit", "dataks"]
ISOLATED_DATE_GAP = 5          # days
MIN_NEIGHBOURS = 3             # fewer other interviews than this within ±5 days → suspicious date
DEFAULT_THRESHOLD = 10000      # € – "shpenzime ≤ 10,000 €"

NUMBER_WORDS = {
    "asnje": 0, "asnjë": 0, "zero": 0, "ska": 0, "s'ka": 0, "nuk ka": 0, "0": 0,
    "nje": 1, "një": 1, "dy": 2, "tre": 3, "tri": 3, "kater": 4, "katër": 4, "pese": 5, "pesë": 5,
    "gjashte": 6, "gjashtë": 6, "shtate": 7, "shtatë": 7, "tete": 8, "tetë": 8, "nente": 9, "nëntë": 9,
    "dhjete": 10, "dhjetë": 10,
}

# Country of residence (shteti): code → (ISO, name). Spellings as in the 2026 summary sheet; others from the codebook.
COUNTRIES = {
    1: ("CH", "Zvicër"), 2: ("DE", "Gjermani"), 3: ("US", "SHBA"), 4: ("SE", "Suedi"), 5: ("GB", "Mbretëria e Bashkuar"),
    6: ("AT", "Austri"), 7: ("FI", "Finlandë"), 8: ("NO", "Norvegji"), 9: ("BE", "Belgjikë"), 10: ("DK", "Danimarkë"),
    11: ("FR", "Francë"), 12: ("NL", "Holandë"), 13: ("IS", "Islandë"), 14: ("HU", "Hungari"), 15: ("IT", "Itali"),
    16: ("CA", "Kanada"), 17: ("HR", "Kroaci"), 18: ("LU", "Luksemburg"), 19: ("PL", "Poloni"), 20: ("SI", "Slloveni"),
    21: ("SK", "Sllovaki"), 22: ("BG", "Bullgari"), 23: ("BA", "Bosnjë dhe Hercegovinë"), 24: ("BY", "Bjellorusi"),
    25: ("CZ", "Çeki"), 26: ("AU", "Australi"), 27: ("RS", "Serbi"), 28: ("TR", "Turqi"), 29: ("GR", "Greqi"),
    30: ("IE", "Irlandë"), 31: ("PT", "Portugali"), 32: ("RO", "Rumani"), 33: ("ES", "Spanjë"), 34: ("CN", "Kinë"),
}
# Countries visited outside Kosovo (vendi)
PLACES = {1: "Shqipëri", 2: "Mal i Zi", 3: "Kroaci", 4: "Spanjë", 5: "Turqi", 6: "Greqi", 7: "Portugali", 8: "Itali"}

# Columns of the 2026 export. Missing optional ones are added empty (with a warning); unknown ones are kept.
KNOWN_COLUMNS = ["interview__key", "interview__id", "referues", "datapytesorit", "pytesorivendi", "anketuesi", "group",
                 "nrpers", "nrpers18", "shteti", "pasaporta", "miksPasaportKosovare", "miksPasaportaHuaj", "Qellimi",
                 "dataks", "vizitejashtks", "vendi", "ditet", "remitanca", "transfer__1", "transfer__2", "transfer__3",
                 "transfer__4", "transfer__5", "totali", "para_arritjes", "akomodimi", "blerjaProd", "Ushqim", "argetimi",
                 "shpenzimetTran", "shpenzimeSh", "DhurataPer", "ParaFalur", "blerjePatun", "shpenzimeTjera",
                 "vleraShpenzimit", "PaguarKartel", "dhuratangaKosova", "tetjera", "shumaShpenzimit", "deposita", "forma",
                 "sssys_irnd", "has__errors", "interview__status", "assignment__id"]
OPTIONAL_COLUMNS = ["interview__id", "nrpers18", "shteti", "pasaporta", "miksPasaportKosovare", "miksPasaportaHuaj",
                    "Qellimi", "vendi", "remitanca", "transfer__1", "transfer__2", "transfer__3", "transfer__4",
                    "transfer__5", "shpenzimeTjera", "PaguarKartel", "dhuratangaKosova", "tetjera", "shumaShpenzimit",
                    "deposita", "forma", "assignment__id"]
# If a year's export names a column differently, map it here: {"old name": "2026 name"}
RENAME = {}

SPEND = ["para_arritjes", "akomodimi", "blerjaProd", "Ushqim", "argetimi", "shpenzimetTran", "shpenzimeSh",
         "DhurataPer", "ParaFalur", "blerjePatun", "vleraShpenzimit"]
BLOCK = ["BR"]                 # column of the indicator block results (set from the data in main)
IN_KOSOVO = ["akomodimi", "blerjaProd", "Ushqim", "argetimi", "shpenzimetTran", "shpenzimeSh", "vleraShpenzimit"]

DERIVED = ["Pika kufitare", "Nr. personave", "Diferenca ne data", "Ditet e qendrimit ne KS", "Total Shpenzimet",
           "Diferenca", "Korrigjimi", "Shpenzimet tjera të paidentifikuara", "shudhetimit", "shmesatare",
           "Total shpz.udh", "shpz2", "Total shpenzimet për person/ditë", "Shpenzimet pa para_arritjes",
           "Shpenzimet pa para_arritjes për person/ditë", "Outlier (lista)", "Total Shpenzimet > 10000",
           "Vlerë ekstreme (shpenzimet pa para_arritjes)", "nrpers18 = nrpers", 'Nr. përgjigjesh "nuk di" (Final)']


# =============================================================================================  helpers
def norm_word(s):
    s = str(s).strip().lower()
    return s


def is_num(v):
    return isinstance(v, (int, float, np.integer, np.floating)) and not (isinstance(v, float) and np.isnan(v))


def nz(v):
    """Excel N(): numbers stay, everything else → 0."""
    return float(v) if is_num(v) else 0.0


class Log:
    def __init__(self):
        self.rows = []

    def add(self, key, col, old, new, rule):
        self.rows.append((key, col, "" if old is None else str(old), "" if new is None else str(new), rule))


# =============================================================================================  0. year schema
# Everything that differs between years (column names, codes, labels) is read from skema.json and the codebook and
# the export is translated into the 2026 layout first. The rest of the script then works the same for every year.
SCHEMA = {}
PIKAT = {1: ("M", "Merdarë"), 2: ("A", "Aeroport")}   # code in pytesorivendi → (letter, name in tables)
ABSENT = set()                                         # columns not in this year's export (tables marked)


def norm_key(s):
    """'Vizite jashtë Kosovës' → 'vizitejashtekosoves' (no case, diacritics, spaces or signs)."""
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9_]", "", s)


def load_schema(path):
    global ISOLATED_DATE_GAP, MIN_NEIGHBOURS, DEFAULT_THRESHOLD
    with open(path, encoding="utf-8") as f:
        S = json.load(f)
    SCHEMA.clear(); SCHEMA.update(S)
    MISSING_CODES[:] = S.get("kodet_e_munguara", MISSING_CODES)
    NA_STRINGS[:] = S.get("tekstet_e_munguara", NA_STRINGS)
    TEXT_COLUMNS[:] = S.get("kolona_teksti", TEXT_COLUMNS)
    DATE_COLUMNS[:] = S.get("kolona_date", DATE_COLUMNS)
    if "numra_me_fjale" in S:
        NUMBER_WORDS.clear(); NUMBER_WORDS.update(S["numra_me_fjale"]); NUMBER_WORDS["0"] = 0
    ISOLATED_DATE_GAP = S.get("data", {}).get("dite_dalluese", ISOLATED_DATE_GAP)
    MIN_NEIGHBOURS = S.get("data", {}).get("min_intervista_fqinje", MIN_NEIGHBOURS)
    DEFAULT_THRESHOLD = S.get("pragu_shpenzimeve", DEFAULT_THRESHOLD)
    KNOWN_COLUMNS[:] = [c for c in S["kolonat"] if not c.startswith("_")]
    return S


# ---- what each column should contain (used when a header is missing, misspelled or shifted)
def _missing(x):
    return (x is None or (isinstance(x, float) and np.isnan(x)) or (is_num(x) and x in MISSING_CODES)
            or (isinstance(x, str) and (x.strip() == "" or x.strip() in NA_STRINGS)))


def _num(x):
    if is_num(x):
        return float(x)
    if isinstance(x, str):
        t = norm_word(x)
        if t in NUMBER_WORDS:
            return float(NUMBER_WORDS[t])
        try:
            return float(t.replace(",", "."))
        except ValueError:
            return None
    return None


def _date(x):
    if isinstance(x, (datetime, pd.Timestamp)):
        return 1990 <= x.year <= 2100
    if is_num(x):
        return 32874 < x < 73051                      # Excel serial numbers 1990–2100
    if isinstance(x, str) and re.fullmatch(r"\s*\d{1,4}[./-]\d{1,2}[./-]\d{1,4}.*", x):
        t = pd.to_datetime(x.strip(), errors="coerce", dayfirst=True)
        return not pd.isna(t) and 1990 <= t.year <= 2100
    return False


def _codes(nums, words=()):
    def f(x):
        v = _num(x)
        if v is not None and v == int(v) and int(v) in nums:
            return True
        return isinstance(x, str) and norm_key(x) in words
    return f


def _rng(lo, hi, integer=False):
    def f(x):
        v = _num(x)
        return v is not None and lo <= v <= hi and (not integer or v == int(v))
    return f


_YESNO = ("po", "jo")
_PIKA_TXT = ("a", "m", "d", "1a", "2m", "aeroport", "aeroporti", "airport", "merdare")
CHECKS = {
    "interview__key": lambda x: isinstance(x, str) and bool(re.fullmatch(r"\d{2}-\d{2}-\d{2}-\d{2}", x.strip())),
    "interview__id": lambda x: isinstance(x, str) and bool(re.fullmatch(r"[0-9a-fA-F]{16,}", x.strip())),
    "datapytesorit": _date, "dataks": _date,
    "pytesorivendi": _codes({1, 2, 3}, _PIKA_TXT),
    "anketuesi": _rng(1, 999, True),
    "group": _codes({1, 2}), "vizitejashtks": _codes({1, 2}), "Qellimi": _codes({1, 2}),
    "nrpers": _rng(1, 500, True), "nrpers18": _rng(0, 500, True), "ditet": _rng(0, 400, True),
    "miksPasaportKosovare": _rng(0, 500, True), "miksPasaportaHuaj": _rng(0, 500, True),
    "pasaporta": _codes({1, 2, 3}), "forma": _codes({1, 2, 3}),
    "shteti": lambda x: _rng(1, 300, True)(x) or (isinstance(x, str) and x.strip().isalpha()),
    "vendi": lambda x: _rng(1, 50, True)(x) or (isinstance(x, str) and x.strip().isalpha()),
    "remitanca": _codes({0, 1, 2}, _YESNO), "deposita": _codes({0, 1, 2}, _YESNO),
    **{f"transfer__{k}": lambda x: _rng(0, 100)(x) or (isinstance(x, str) and norm_key(x) in _YESNO) for k in range(1, 6)},
    **{c: _rng(0, 1e7) for c in SPEND + ["totali", "dhuratangaKosova", "shumaShpenzimit"]},
}


def col_score(name, series):
    """Share of the filled cells that look like `name` should (None = no rule or empty column)."""
    f = CHECKS.get(name)
    if f is None:
        return None
    v = [x for x in series.head(400) if not _missing(x)]
    if not v:
        return None
    share = sum(1 for x in v if f(x)) / len(v)
    distinct = len(set(map(str, v)))
    if name == "anketuesi" and distinct < 3:            # several enumerators, unlike a 1/2 code
        share *= 0.3
    if name == "referues" and distinct < 0.9 * len(v):  # running number: (almost) all different
        share *= 0.3
    if name == "pytesorivendi" and distinct > 4:
        share *= 0.3
    return share


def _match_names(cols, spec, raw=None):
    """{position: canonical name} from the headers: exact name, aliases, then close spellings."""
    import difflib
    pos = {}
    for i, c in enumerate(cols):
        if c in spec and c not in pos.values():
            pos[i] = c
    keys = {i: norm_key(c) for i, c in enumerate(cols)}
    for c, aliases in spec.items():
        if c in pos.values():
            continue
        for a in aliases:
            hit = next((i for i in keys if i not in pos and keys[i] == norm_key(a)), None)
            if hit is not None:
                pos[hit] = c
                break
    for c, aliases in spec.items():                     # 'Pika kufitare (1=M,2=A)', 'Datapytsorit', …
        if c in pos.values():
            continue
        best = None
        for i, k in keys.items():
            if i in pos or len(k) < 4:
                continue
            r = max([difflib.SequenceMatcher(None, k, norm_key(a)).ratio() for a in aliases] +
                    [1.0 if any(len(norm_key(a)) >= 5 and norm_key(a) in k for a in aliases) else 0.0])
            sc = col_score(c, raw.iloc[:, i]) if raw is not None and i < raw.shape[1] else None
            if r >= 0.85 and (sc is None or sc >= 0.8) and (best is None or r > best[0]):
                best = (r, i)
        if best:
            pos[best[1]] = c
    return pos


def _align(raw, names):
    """Header row shifted against the data (a cell deleted or inserted in the header)? Pick, for every header, the
    column offset (−1, 0, +1) that best fits the contents, with a penalty for changing offset (dynamic programming)."""
    n = len(names)
    offs = (-1, 0, 1)
    cache = {}

    def cost(i, o):
        c = names[i]
        if c is None:
            return 0.0
        j = i + o
        if not 0 <= j < raw.shape[1]:
            return 9.0
        if (c, j) not in cache:
            cache[(c, j)] = col_score(c, raw.iloc[:, j])
        s = cache[(c, j)]
        return (1.0 - (0.5 if s is None else s)) + 0.02 * abs(o)

    INF = float("inf")
    best = [{o: (cost(0, o), [o]) for o in offs}]
    for i in range(1, n):
        row = {}
        for o in offs:
            cands = [(best[-1][p][0] + (0 if p == o else 1.0), best[-1][p][1]) for p in offs if p <= o]
            c, path = min(cands, key=lambda t: t[0]) if cands else (INF, [])
            row[o] = (c + cost(i, o), path + [o])
        best.append(row)
    total, path = min(best[-1].values(), key=lambda t: t[0])
    stay = sum(cost(i, 0) for i in range(n))
    return path if total < stay - 0.5 else [0] * n


def harmonise_columns(raw, log):
    """Find every column of the 2026 layout in this year's export: by name (exact, alias, close spelling), by its
    contents when the header row is shifted, and by position when a header is empty. Nothing is fatal: a column that
    cannot be found is added empty and the tables that need it are marked."""
    spec = {c: a for c, a in SCHEMA["kolonat"].items() if not c.startswith("_")}
    # 1. header row: the first row of the sheet is not always the header (a title above it)
    pos = _match_names(list(raw.columns), spec, raw)
    if len(pos) < 0.3 * len(spec):
        for r in range(min(10, len(raw))):
            hdr = [str(x) if not _missing(x) else f"Unnamed: {k}" for k, x in enumerate(raw.iloc[r])]
            p2 = _match_names(hdr, spec, raw.iloc[r + 1:])
            if len(p2) > len(pos):
                raw = raw.iloc[r + 1:].reset_index(drop=True)
                raw.columns = hdr
                pos = p2
                log.add("—", "—", "", f"rreshti {r + 2}", "KUJDES: titujt e kolonave nuk ishin në rreshtin e parë – u gjetën më poshtë")
        raw = raw.dropna(how="all").reset_index(drop=True)
    cols = list(raw.columns)
    names = [pos.get(i) for i in range(len(cols))]
    for i, c in enumerate(names):
        if c and cols[i] != c:
            log.add("—", c, cols[i], c, "Kolona u njoh nga emri (emër i vjetër ose i shkruar ndryshe) dhe u riemërtua")
    # 2. header shifted against the data → follow the contents
    path = _align(raw, names)
    if any(path):
        moved = [f"{names[i]}: kolona {xl_col_to_name(i)} → {xl_col_to_name(i + o)}" for i, o in enumerate(path) if o and names[i]]
        log.add("—", "—", "", "", f"KUJDES: titujt e kolonave ishin të zhvendosur nga të dhënat – {len(moved)} kolona u "
                "lidhën sipas përmbajtjes (" + "; ".join(moved[:3]) + ("; …" if len(moved) > 3 else "") + ")")
    where = {names[i]: i + o for i, o in enumerate(path) if names[i] and 0 <= i + o < len(cols)}
    # 3. a header whose column clearly does not fit it (e.g. two headers swapped) → swap or drop it
    sc = lambda c, j: col_score(c, raw.iloc[:, j])
    for c in list(where):
        if c not in where:
            continue
        j = where[c]
        s0 = sc(c, j)
        if s0 is None or s0 >= 0.5:
            continue
        partner = next((d for d, k in where.items() if d != c and abs(k - j) <= 3
                        and (sc(c, k) or 0) >= 0.9 and (sc(d, j) or 0) >= 0.9), None)
        if partner:
            where[c], where[partner] = where[partner], j
            log.add("—", c, xl_col_to_name(j), xl_col_to_name(where[c]), f"KUJDES: titujt '{c}' dhe '{partner}' ishin "
                    "të ngatërruar – u këmbyen sipas përmbajtjes")
        else:
            del where[c]
            log.add("—", c, str(cols[j]), "", f"KUJDES: kolona me titullin '{cols[j]}' nuk përmban atë që pritet për '{c}' "
                    f"(vetëm {s0:.0%} e vlerave përshtaten) – nuk u përdor")
    # 4. a column whose header is empty/unknown, sitting where a missing column should be
    order = [c for c in spec]
    used = set(where.values())
    for k, c in enumerate(order):
        if c in where:
            continue
        before = [where[x] for x in order[:k] if x in where]
        after = [where[x] for x in order[k + 1:] if x in where]
        lo, hi = (max(before) if before else -1), (min(after) if after else len(cols))
        cands = [j for j in range(lo + 1, hi) if j not in used]
        scored = [(col_score(c, raw.iloc[:, j]) or 0, j) for j in cands]
        good = [t for t in scored if t[0] >= 0.9]
        if len(good) == 1:
            where[c] = good[0][1]; used.add(good[0][1])
            log.add("—", c, str(cols[good[0][1]]), c, f"KUJDES: kolona pa titull/me titull të panjohur në vendin e '{c}' "
                    "u njoh sipas pozitës dhe përmbajtjes – kontrollo")
    # build the frame in the export's column order
    out = pd.DataFrame(index=raw.index)
    inv = {j: c for c, j in where.items()}
    for j, col in enumerate(cols):
        if j in inv:
            out[inv[j]] = raw.iloc[:, j]
        else:
            nm = str(col)
            if nm in spec or nm in out.columns or nm.startswith("Unnamed"):
                if raw.iloc[:, j].map(_missing).all():
                    continue                              # empty column without a header
                nm = f"kolona_{xl_col_to_name(j)}"
            out[nm] = raw.iloc[:, j]
    raw = out
    found = [c for c in spec if c in raw.columns]
    if len(found) < 0.4 * len(spec):
        raise SystemExit("Ky skedar nuk duket si eksporti i anketës: u gjetën vetëm " + str(len(found)) + " nga "
                         + str(len(spec)) + " kolonat (" + ", ".join(found) + "). Kontrolloni skedarin/fletën "
                         "(--fleta) ose shtoni emrat e kolonave te 'kolonat' në skema.json.")
    # identifiers that older exports (before Survey Solutions) do not have
    if "referues" not in raw.columns:
        raw.insert(0, "referues", range(1, len(raw) + 1))
        log.add("—", "referues", "", "1…n", "KUJDES: kolona e numrit rendor mungon – u krijua sipas radhës")
    if "interview__key" not in raw.columns:
        raw.insert(0, "interview__key", ["R" + str(int(v)) if is_num(v) else f"R{i + 1}" for i, v in enumerate(raw["referues"])])
        log.add("—", "interview__key", "", "R<referues>", "KUJDES: kolona interview__key mungon – u krijua nga referues")
    quiet = set(SCHEMA.get("kolona_pa_tabela", []))
    must = set(SCHEMA.get("te_domosdoshme", []))
    for c in KNOWN_COLUMNS:
        if c not in raw.columns:
            raw[c] = np.nan
            ABSENT.add(c)
            log.add("—", c, "", "", "Kolona mungon në eksport – u shtua bosh (nuk përdoret në tabela)" if c in quiet
                    else ("KUJDES: kolonë e rëndësishme mungon në eksport – u shtua bosh; rezultatet që varen prej saj "
                          "dalin bosh/0, tabelat shënohen" if c in must else
                          "KUJDES: kolona mungon në eksport – u shtua bosh, tabelat përkatëse shënohen"))
    for c in raw.columns:
        if c not in KNOWN_COLUMNS:
            log.add("—", c, "", "", "Kolonë e panjohur në eksport – u ruajt, por nuk përdoret në tabela")
    return raw


def guess_year(raw):
    d = pd.to_datetime(raw["datapytesorit"].map(to_date), errors="coerce")
    return int(d.dt.year.mode().iloc[0]) if d.notna().any() else datetime.now().year


def read_codebook(path, year):
    """Sections of the codebook sheet for `year` (or the latest earlier one): {column: [rows]}."""
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    sheets = {}
    for s in wb.sheetnames:
        m = re.search(r"(19|20)\d\d", s)
        if m:
            sheets[int(m.group(0))] = s
    if not sheets:
        sheets = {year: wb.sheetnames[0]}
    pick = year if year in sheets else max([y for y in sheets if y <= year] or [min(sheets)])
    alias = {norm_key(a): c for c, al in SCHEMA["kolonat"].items() if not c.startswith("_") for a in al}
    sec, out = None, {}
    for r in wb[sheets[pick]].iter_rows(values_only=True):
        v = [x for x in r if x is not None and str(x).strip() != ""]
        if not v:
            continue
        qnum = re.fullmatch(r"\d+[a-z]*", str(v[0]).strip())
        if len(v) == 2 and isinstance(v[1], str) and qnum and norm_key(v[1]) in alias:
            sec = alias[norm_key(v[1])]; out[sec] = []
            continue
        if len(v) == 2 and isinstance(v[1], str) and qnum and not str(v[0]).strip().isdigit():
            sec = None                                  # a question we do not use (e.g. '10a')
            continue
        if sec and str(v[0]).strip().lower() != "label":
            out[sec].append(v)
    return sheets[pick], out


def set_year_codes(year, codebook, log):
    """Code tables for this year: border points, countries of residence, countries visited."""
    S = SCHEMA
    names = S.get("emrat_e_shteteve", {})
    code_iso = {int(k): v for k, v in S.get("kodet_e_shteteve_parazgjedhje", {}).items()}
    places = {int(k): v for k, v in S.get("kodet_e_vendeve_parazgjedhje", {}).items()}
    place_names = {k: names.get(v, v) for k, v in places.items()}
    pk = S.get("pika_kufitare", {})
    pika_num = {int(k): v for k, v in pk.get(str(year), pk.get("parazgjedhje", {"1": "M", "2": "A"})).items()}
    src = f"skema.json ({'viti ' + str(year) if str(year) in pk else 'parazgjedhje'})"
    extra_names = {}
    if codebook:
        sheet, cb = read_codebook(codebook, year)
        log.add("—", "—", "", sheet, "Codebook-u i përdorur për kodet e këtij viti")
        num = {}
        for v in cb.get("pytesorivendi", []):
            if len(v) >= 2 and isinstance(v[1], str):
                lbl = str(v[0]).strip().upper()
                m = re.fullmatch(r"(\d+)\s*([A-Z])", lbl)
                if m:
                    num[int(m.group(1))] = m.group(2)
                    extra_names[m.group(2)] = v[1].strip()
                elif re.fullmatch(r"[A-Z]", lbl):
                    extra_names[lbl] = v[1].strip()
        if num:
            pika_num, src = num, f"codebook ({sheet})"
        cc = {v[0]: str(v[1]).strip().upper() for v in cb.get("shteti", [])
              if len(v) >= 3 and isinstance(v[0], int) and isinstance(v[1], str) and len(v[1].strip()) == 2}
        if cc:
            code_iso = cc
            for v in cb["shteti"]:
                if len(v) >= 3 and isinstance(v[0], int) and isinstance(v[1], str):
                    names.setdefault(v[1].strip().upper(), str(v[2]).strip())
        pl, pn = {}, {}
        no_iso = {norm_key(k): v for k, v in S.get("emrat_pa_iso", {}).items() if not k.startswith("_")}
        for v in cb.get("vendi", []):
            if not isinstance(v[0], int) or len(v) < 2:
                continue
            if len(v) >= 3 and isinstance(v[1], str) and len(v[1].strip()) == 2:
                iso, nm = v[1].strip().upper(), str(v[2]).strip()
            else:
                nm = str(v[1]).strip(); iso = no_iso.get(norm_key(nm))
            pl[v[0]] = iso
            pn[v[0]] = names.get(iso, nm) if iso else nm
        if pl:
            places, place_names = pl, pn
    COUNTRIES.clear()
    COUNTRIES.update({c: (iso, names.get(iso, iso)) for c, iso in code_iso.items()})
    PLACES.clear(); PLACES.update(place_names)
    PIKAT.clear(); PIKAT.update({1: ("M", "Merdarë"), 2: ("A", "Aeroport")})
    log.add("—", "pytesorivendi", "", ", ".join(f"{k}={v}" for k, v in sorted(pika_num.items())),
            "Kuptimi i kodeve numerike të pikës kufitare, nga " + src)
    return pika_num, extra_names


def harmonise_values(df, year, pika_num, extra_names, log):
    """Translate this year's codes into the 2026 codes (Pika 1=M/2=A, Po=1/Jo=2, numeric country codes, 0/1 transfers)."""
    keys = df["interview__key"].astype(str)
    S = SCHEMA
    tekste = {norm_key(k): v for k, v in S.get("pika_kufitare", {}).get("tekste", {}).items()}
    # --- border point → letter → 2026 code
    def letter(v):
        if is_num(v) and v in MISSING_CODES:
            return v
        if is_num(v):
            return pika_num.get(int(v))
        k = norm_key(v)
        if k in tekste:
            return tekste[k]
        m = re.fullmatch(r"(\d+)([a-z])", k)
        if m:
            return m.group(2).upper()
        if k.isdigit():
            return pika_num.get(int(k))
        return k.upper() if len(k) == 1 else None
    let = df["pytesorivendi"].map(lambda v: None if (v is None or (isinstance(v, float) and np.isnan(v))) else letter(v))
    code_of = {"M": 1, "A": 2}
    for L_ in sorted({x for x in let.dropna() if isinstance(x, str)} - {"M", "A"}):
        code_of[L_] = max(code_of.values()) + 1
        PIKAT[code_of[L_]] = (L_, extra_names.get(L_, f"Pika {L_}"))
        log.add("—", "pytesorivendi", L_, code_of[L_], f"KUJDES: pikë kufitare tjetër ({PIKAT[code_of[L_]][1]}) – "
                "numërohet në DB dhe ka fletën e vet, por tabelat A/M nuk e përfshijnë")
    new = let.map(lambda x: x if is_num(x) else (code_of.get(x, np.nan) if x else np.nan))
    oldn = pd.to_numeric(df["pytesorivendi"], errors="coerce")
    changed = ~((oldn == new) | (oldn.isna() & new.isna()))
    if changed.any():
        for (o, n_), cnt in pd.Series(list(zip(df["pytesorivendi"][changed].astype(str), new[changed]))).value_counts().items():
            log.add("—", "pytesorivendi", o, "" if pd.isna(n_) else int(n_),
                    f"Kodi i pikës u kthye në kodin e 2026 (1=M, 2=A) – {cnt} intervista")
    df["pytesorivendi"] = new
    # --- yes/no questions: 'Po'/'Jo' → 1/2, and 0 = Jo → 2
    for c in S.get("po_jo", {}).get("kolonat", []):
        if c in ABSENT:
            continue
        col = df[c].map(lambda v: {"po": 1, "jo": 2}.get(norm_key(v), v) if isinstance(v, str) else v)
        vals = set(pd.to_numeric(col, errors="coerce").dropna().unique()) - set(MISSING_CODES)
        if 0 in vals and vals <= {0, 1}:
            col = col.map(lambda v: 2 if is_num(v) and v == 0 else v)
            log.add("—", c, "0", "2", f"Jo ishte 0 në këtë vit → 2 (si në 2026) – {int((df[c] == 0).sum())} intervista")
        n_txt = int(df[c].map(lambda v: isinstance(v, str) and norm_key(v) in ("po", "jo")).sum())
        if n_txt:
            log.add("—", c, "Po/Jo", "1/2", f"Përgjigje me tekst → kod – {n_txt} intervista")
        df[c] = col
    # --- remittance channels: 0/1 in 2026; percentages or Po/Jo in some years
    for k in range(1, 6):
        c = f"transfer__{k}"
        if c in ABSENT:
            continue
        col = df[c].map(lambda v: {"po": 1, "jo": 0}.get(norm_key(v), v) if isinstance(v, str) else v)
        x = pd.to_numeric(col, errors="coerce")
        if (x > 1).any():
            log.add("—", c, "përqindje", "1/0", f"Kanali i remitancave ishte në përqindje → 1 nëse > 0 – "
                    f"{int((x > 0).sum())} intervista me 1")
            col = x.map(lambda v: np.nan if pd.isna(v) else (v if v in MISSING_CODES else (1 if v > 0 else 0)))
        df[c] = col
    # --- countries (residence) and countries visited: text/ISO → numeric code of this year
    for c, table, add in (("shteti", COUNTRIES, True), ("vendi", PLACES, False)):
        if c in ABSENT:
            continue
        if c == "shteti":
            by_txt = {iso.upper(): code for code, (iso, _) in table.items()}
            by_txt.update({norm_key(nm): code for code, (_, nm) in table.items()})
        else:
            by_txt = {norm_key(nm): code for code, nm in table.items()}
            by_txt.update({iso: code for code, iso in
                           {int(k): v for k, v in SCHEMA.get("kodet_e_vendeve_parazgjedhje", {}).items()}.items()})
        names = SCHEMA.get("emrat_e_shteteve", {})
        recoded = {}
        def to_code(v):
            if not isinstance(v, str) or v.strip() in NA_STRINGS:
                return v
            t = v.strip()
            if re.fullmatch(r"\d+", t):
                return int(t)
            hit = by_txt.get(t.upper()) if len(t) == 2 else by_txt.get(norm_key(t))
            if hit is None and len(t) == 2 and t.isalpha():
                hit = max(list(table) + [99]) + 1          # a country not in this year's list
                table[hit] = (t.upper(), names.get(t.upper(), t.upper())) if add else names.get(t.upper(), t.upper())
                log.add("—", c, t, hit, "KUJDES: shtet pa kod në listën e vitit – iu dha kod i ri")
            if hit is not None:
                recoded[t] = hit
                return hit
            return v
        df[c] = df[c].map(to_code)
        for t, h in recoded.items():
            log.add("—", c, t, h, "Shteti me tekst → kod numerik")
    return df


def sanity_checks(db, log):
    """Warnings that point to a wrong code table (not errors)."""
    pa = db.groupby("Pika kufitare")["para_arritjes"].mean()
    if "A" in pa and "M" in pa and pa["A"] < pa["M"]:
        log.add("—", "pytesorivendi", "", "", "KUJDES: shpenzimet para mbërritjes janë më të ulëta në Aeroport se në "
                f"Merdarë ({pa['A']:.0f} € kundër {pa['M']:.0f} €) – kontrolloni a janë këmbyer kodet e pikës (skema.json)")
    for c, table in (("shteti", COUNTRIES), ("vendi", PLACES)):
        unk = sorted(set(pd.to_numeric(db[c], errors="coerce").dropna().astype(int)) - set(table))
        if unk:
            log.add("—", c, "", "", f"KUJDES: kode pa etiketë {unk} – jepni codebook-un e vitit (--codebook)")


# =============================================================================================  1. cleaning
def to_date(v):
    """Excel date, Excel serial number or text (day first: 05/08/2018 = 5 August)."""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return pd.NaT
    if isinstance(v, (datetime, pd.Timestamp)):
        return pd.Timestamp(v)
    if is_num(v):
        return pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(v)) if 20000 < v < 80000 else pd.NaT
    return pd.to_datetime(str(v).strip(), errors="coerce", dayfirst=True)


def clean(raw, log, korrigjime=None):
    df = raw.copy()
    keys = df["interview__key"].astype(str)
    if "_nukdi" not in df:
        df["_nukdi"] = (df.isin(MISSING_CODES)).sum(axis=1)
    for c in df.columns:
        if c in ("interview__key", "_nukdi"):
            continue
        col = df[c].astype(object)
        for i, v in col.items():
            if is_num(v) and v in MISSING_CODES:
                col[i] = np.nan
                log.add(keys[i], c, v, "", "Kod për 'nuk e di/nuk përgjigjet' → bosh")
            elif isinstance(v, str) and v.strip() in NA_STRINGS:
                col[i] = np.nan
            elif isinstance(v, str) and c not in TEXT_COLUMNS and c not in DATE_COLUMNS:
                w = norm_word(v)
                if w in NUMBER_WORDS:
                    col[i] = NUMBER_WORDS[w]
                    log.add(keys[i], c, v, NUMBER_WORDS[w], "Numër i shkruar me fjalë → numër")
                else:
                    try:
                        col[i] = float(w.replace(",", "."))
                        log.add(keys[i], c, v, col[i], "Numër si tekst → numër")
                    except ValueError:
                        col[i] = np.nan
                        log.add(keys[i], c, v, "", "Tekst i panjohur në fushë numerike → bosh (kontrollo)")
        if c in DATE_COLUMNS:
            df[c] = pd.to_datetime(col.map(to_date))
        elif c in TEXT_COLUMNS:
            df[c] = col
        else:
            df[c] = pd.to_numeric(col, errors="coerce")
    fix_isolated_dates(df, log)
    if korrigjime:
        apply_corrections(df, korrigjime, log)
    return df


def fix_isolated_dates(df, log):
    """An interview date with fewer than MIN_NEIGHBOURS other interviews within ±ISOLATED_DATE_GAP days is a likely
    typing error (2026: one interview dated 19 July, fieldwork began 27 July). It is replaced with the date of the
    neighbouring interviews of the same assignment (or, in exports without assignment__id, the same enumerator) when
    they agree; otherwise it is only flagged in the Log."""
    d = df["datapytesorit"]
    near = {i: int(((d - v).abs().dt.days <= ISOLATED_DATE_GAP).sum()) - 1 for i, v in d.items() if not pd.isna(v)}
    suspicious = {i for i, k in near.items() if k < MIN_NEIGHBOURS}
    for i in sorted(suspicious):
        v = d[i]
        key = df.at[i, "interview__key"]
        others = d.drop(i).dropna()
        gap = int((others - v).abs().min().days) if len(others) else 0
        pos = df.index.get_loc(i)
        grp = "assignment__id" if "assignment__id" in df and df["assignment__id"].notna().any() else "anketuesi"
        same = [j for j in df.index if j != i and j not in suspicious and df.at[j, grp] == df.at[i, grp]]
        prev = [j for j in same if df.index.get_loc(j) < pos]
        nxt = [j for j in same if df.index.get_loc(j) > pos]
        cand = ([d[prev[-1]]] if prev else []) + ([d[nxt[0]]] if nxt else [])
        if cand and all(c == cand[0] for c in cand) and not pd.isna(cand[0]) and cand[0] != v:
            df.at[i, "datapytesorit"] = cand[0]
            log.add(key, "datapytesorit", v.date(), cand[0].date(),
                    f"Datë e izoluar ({gap} ditë larg intervistave tjera) → data e intervistave fqinje të të njëjtit "
                    + ("assignment" if grp == "assignment__id" else "anketues"))
        else:
            log.add(key, "datapytesorit", v.date(), v.date(),
                    f"KUJDES: datë e izoluar ({gap} ditë larg intervistave tjera) – nuk u korrigjua automatikisht")


def apply_corrections(df, path, log):
    with open(path, newline="", encoding="utf-8") as f:
        for r in csv.DictReader(line for line in f if not line.lstrip().startswith("#")):
            m = df["interview__key"].astype(str) == r["interview__key"].strip()
            if not m.any():
                log.add(r["interview__key"], r["kolona"], "", r["vlera"], "KUJDES: korrigjim për intervistë që nuk ekziston")
                continue
            col = r["kolona"].strip()
            old = df.loc[m, col].iloc[0]
            new = pd.to_datetime(r["vlera"]) if col in DATE_COLUMNS else (
                r["vlera"] if col in TEXT_COLUMNS else pd.to_numeric(r["vlera"], errors="coerce"))
            df.loc[m, col] = new
            log.add(r["interview__key"], col, old, new, "Korrigjim manual: " + r.get("arsyeja", ""))


# =============================================================================================  2. derived columns
def derive(df, outliers, threshold):
    o = pd.DataFrame(index=df.index)
    E, G, H, I, P, R, Y = (df[c] for c in ["pytesorivendi", "group", "nrpers", "nrpers18", "vizitejashtks", "ditet", "totali"])
    o["Pika kufitare"] = E.map(lambda v: PIKAT[int(v)][0] if is_num(v) and int(v) in PIKAT else "")
    o["Nr. personave"] = np.where(G == 1, 1, H.fillna(0))
    o["Diferenca ne data"] = (df["datapytesorit"] - df["dataks"]).dt.days
    o["Ditet e qendrimit ne KS"] = np.where(P == 1, o["Diferenca ne data"],
                                            np.where((P == 2) & R.notna(), o["Diferenca ne data"] - R, np.nan))
    s = lambda cols: df[cols].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1)
    o["Total Shpenzimet"] = s(SPEND)
    o["Diferenca"] = np.where(Y.notna(), Y - o["Total Shpenzimet"], np.nan)
    o["Korrigjimi"] = o["Diferenca"]
    o["Shpenzimet tjera të paidentifikuara"] = np.where(o["Korrigjimi"].notna(), np.maximum(o["Korrigjimi"], 0), np.nan)
    o["shudhetimit"] = s(IN_KOSOVO) + o["Shpenzimet tjera të paidentifikuara"].fillna(0)
    ay = o["Ditet e qendrimit ne KS"]
    pos = ay.fillna(0) > 0
    aw = o["Nr. personave"]
    o["shmesatare"] = np.where(pos, o["shudhetimit"] / aw / ay, np.nan)
    o["Total shpz.udh"] = np.where(Y.notna(), Y - s(["para_arritjes", "DhurataPer", "ParaFalur", "blerjePatun", "vleraShpenzimit"]), np.nan)
    o["shpz2"] = np.where(pos & o["Total shpz.udh"].notna(), o["Total shpz.udh"] / aw / ay, np.nan)
    o["Total shpenzimet për person/ditë"] = np.where(pos, o["Total Shpenzimet"] / ay / aw, np.nan)
    o["Shpenzimet pa para_arritjes"] = o["Total Shpenzimet"] - df["para_arritjes"].fillna(0)
    o["Shpenzimet pa para_arritjes për person/ditë"] = np.where(pos, o["Shpenzimet pa para_arritjes"] / aw / ay, np.nan)
    o["Outlier (lista)"] = df["referues"].isin(outliers).astype(int)
    o["Total Shpenzimet > 10000"] = (o["Total Shpenzimet"] > threshold).astype(int)
    o["nrpers18 = nrpers"] = ((G == 2) & I.notna() & (I == H)).astype(int)
    o['Nr. përgjigjesh "nuk di" (Final)'] = df["_nukdi"]
    return o


def exclusion_reason(row):
    d, ay = row["Diferenca ne data"], row["Ditet e qendrimit ne KS"]
    if pd.isna(ay) or ay > 0:
        return None
    if d < 0:
        return "Data e ardhjes pas datës së pyetësorit (datë e pamundur)"
    if ay < 0:
        r = row["ditet"]
        return f"Ditë qëndrimi negative ({int(d)} ditë nga ardhja, {int(r)} ditë jashtë Kosovës) - e pamundur"
    return ("0 ditë qëndrimi; shpenzimet (p.sh. akomodim) tregojnë qëndrim më të gjatë, ndaj data e ardhjes ka gjasë "
            "të jetë gabim - nuk numërohet si 1 ditë")


def quartile_inc(values, q):
    v = np.sort(np.asarray([x for x in values if is_num(x)], dtype=float))
    return float(np.percentile(v, q * 100)) if len(v) else np.nan


# =============================================================================================  3. workbook
class Book:
    def __init__(self, path):
        self.wb = xlsxwriter.Workbook(path, {"nan_inf_to_errors": True})
        f = self.wb.add_format
        self.f = {
            "title": f({"bold": True, "font_size": 14, "font_color": "#1F3864"}),
            "h1": f({"bold": True, "font_size": 12, "font_color": "#1F3864"}),
            "tab": f({"bold": True}),
            "hdr": f({"bold": True, "bg_color": "#D9D9D9", "border": 1}),
            "lbl": f({"border": 1}),
            "tot": f({"bold": True, "border": 1}),
            "int": f({"border": 1, "num_format": "#,##0"}),
            "inttot": f({"bold": True, "border": 1, "num_format": "#,##0"}),
            "dec": f({"border": 1, "num_format": "#,##0.00"}),
            "pct": f({"border": 1, "num_format": "0.0%"}),
            "pcttot": f({"bold": True, "border": 1, "num_format": "0.0%"}),
            "date": f({"num_format": "yyyy-mm-dd"}),
            "bold": f({"bold": True}),
            "note": f({"italic": True, "font_color": "#808080"}),
        }

    def formula(self, ws, r, c, formula, value, fmt=None):
        if isinstance(value, (float, np.floating)) and (np.isnan(value) or np.isinf(value)):
            value = "#DIV/0!"
        elif isinstance(value, (float, np.floating)) and float(value).is_integer():
            value = int(value)                       # store 1, not 1.0 (as Excel does)
        ws.write_formula(r, c, formula, self.f.get(fmt) if fmt else None, value if value is not None else "")


def letters(columns):
    return {c: xl_col_to_name(i) for i, c in enumerate(columns)}


def write_data_sheet(book, name, data, cols, extra_col=None, db_threshold_ref=None, threshold_ref="Permbledhje!$B$5",
                     outlier_ref="Permbledhje!$B$6:$M$6", ext_limit=None):
    """Write raw + derived columns (derived as formulas with stored values) and the indicator block."""
    ws = book.wb.add_worksheet(name)
    L = letters(cols)
    n = len(data)
    last = n + 1
    hdr_cols = cols + ([extra_col] if extra_col else [None]) + ["Treguesi", "Rezultati"]
    for j, h in enumerate(hdr_cols):
        if h is not None:
            ws.write_string(0, j, h, book.f["bold"])
    raw_cols = [c for c in cols if c not in DERIVED]
    for i, (_, row) in enumerate(data.iterrows()):
        r = i + 1
        x = r + 1                                   # Excel row number
        for c in raw_cols:
            v = row[c]
            j = cols.index(c)
            if c in DATE_COLUMNS and not pd.isna(v):
                ws.write_datetime(r, j, v.to_pydatetime(), book.f["date"])
            elif isinstance(v, str):
                ws.write_string(r, j, v)
            elif is_num(v):
                ws.write_number(r, j, float(v))
        F = lambda c: f"{L[c]}{x}"
        fm = {
            "Pika kufitare": "=" + "".join(f'IF({F("pytesorivendi")}={k},"{PIKAT[k][0]}",' for k in sorted(PIKAT)) + '""' + ")" * len(PIKAT),
            "Nr. personave": f'=IF({F("group")}=1,1,{F("nrpers")})',
            "Diferenca ne data": f'={F("datapytesorit")}-{F("dataks")}',
            "Ditet e qendrimit ne KS": f'=IF({F("vizitejashtks")}=1,{F("Diferenca ne data")},IF(AND({F("vizitejashtks")}=2,ISNUMBER({F("ditet")})),{F("Diferenca ne data")}-{F("ditet")},""))',
            "Total Shpenzimet": "=SUM(" + ",".join(F(c) for c in SPEND) + ")",
            "Diferenca": f'=IF(ISNUMBER({F("totali")}),{F("totali")}-{F("Total Shpenzimet")},"")',
            "Korrigjimi": f'={F("Diferenca")}',
            "Shpenzimet tjera të paidentifikuara": f'=IF(ISNUMBER({F("Korrigjimi")}),MAX({F("Korrigjimi")},0),"")',
            "shudhetimit": "=SUM(" + ",".join(F(c) for c in IN_KOSOVO + ["Shpenzimet tjera të paidentifikuara"]) + ")",
            "shmesatare": f'=IF(N({F("Ditet e qendrimit ne KS")})>0,{F("shudhetimit")}/{F("Nr. personave")}/{F("Ditet e qendrimit ne KS")},"")',
            "Total shpz.udh": f'=IF(ISNUMBER({F("totali")}),{F("totali")}-SUM({F("para_arritjes")},{F("DhurataPer")},{F("ParaFalur")},{F("blerjePatun")},{F("vleraShpenzimit")}),"")',
            "shpz2": f'=IF(AND(N({F("Ditet e qendrimit ne KS")})>0,ISNUMBER({F("Total shpz.udh")})),{F("Total shpz.udh")}/{F("Nr. personave")}/{F("Ditet e qendrimit ne KS")},"")',
            "Total shpenzimet për person/ditë": f'=IF(N({F("Ditet e qendrimit ne KS")})>0,{F("Total Shpenzimet")}/{F("Ditet e qendrimit ne KS")}/{F("Nr. personave")},"")',
            "Shpenzimet pa para_arritjes": f'={F("Total Shpenzimet")}-N({F("para_arritjes")})',
            "Shpenzimet pa para_arritjes për person/ditë": f'=IF(N({F("Ditet e qendrimit ne KS")})>0,{F("Shpenzimet pa para_arritjes")}/{F("Nr. personave")}/{F("Ditet e qendrimit ne KS")},"")',
            "Outlier (lista)": f'=IF(ISNUMBER(MATCH({F("referues")},{outlier_ref},0)),1,0)',
            "Total Shpenzimet > 10000": f'=IF({F("Total Shpenzimet")}>{threshold_ref},1,0)',
            "Vlerë ekstreme (shpenzimet pa para_arritjes)": f'=IF(AND(ISNUMBER({F("Shpenzimet pa para_arritjes për person/ditë")}),{F("Shpenzimet pa para_arritjes për person/ditë")}>DB!${BLOCK[0]}$22),1,0)',
            "nrpers18 = nrpers": f'=IF(AND({F("group")}=2,ISNUMBER({F("nrpers18")}),{F("nrpers18")}={F("nrpers")}),1,0)',
        }
        for c, formula in fm.items():
            v = row[c]
            if c == "Pika kufitare":
                book.formula(ws, r, cols.index(c), formula, v)
            else:
                book.formula(ws, r, cols.index(c), formula, "" if pd.isna(v) else float(v))
        ws.write_number(r, cols.index('Nr. përgjigjesh "nuk di" (Final)'), float(row['Nr. përgjigjesh "nuk di" (Final)']))
        if extra_col:
            ws.write_string(r, len(cols), str(row[extra_col]))
    if extra_col:                                    # Perjashtuar has no indicator block
        return ws
    # ---------------------------------------------- indicator block (columns BQ:BR in 2026)
    tq, tr = len(cols) + 1, len(cols) + 2
    rng = lambda c: f"{L[c]}2:{L[c]}{last}"
    BJ, AY, AW = rng("Shpenzimet pa para_arritjes për person/ditë"), rng("Ditet e qendrimit ne KS"), rng("Nr. personave")
    BL, BM = rng("Total Shpenzimet > 10000"), rng("Vlerë ekstreme (shpenzimet pa para_arritjes)")
    d = data

    def wavg(col, mask=None):
        v = pd.to_numeric(d[col], errors="coerce")
        m = v.notna() if mask is None else v.notna() & mask
        w = d["Ditet e qendrimit ne KS"][m] * d["Nr. personave"][m]
        return (v[m] * w).sum() / w.sum() if w.sum() else np.nan

    bj = "Shpenzimet pa para_arritjes për person/ditë"
    m10 = d["Total Shpenzimet > 10000"] == 0
    mx = d["Vlerë ekstreme (shpenzimet pa para_arritjes)"] == 0
    q1 = quartile_inc(d[bj], 0.25)
    q3 = quartile_inc(d[bj], 0.75)
    col_of = lambda c: rng(c)
    block = [
        ("SHPENZIMET PËR PERSON NË DITË (pa shpenzimet para mbërritjes)", None, None),
        ("Të gjitha intervistat", f"=SUMPRODUCT({BJ},{AY},{AW})/SUMPRODUCT(--ISNUMBER({BJ}),{AY},{AW})", wavg(bj)),
        ("Vetëm intervistat me shpenzime ≤ 10,000 €", f"=SUMPRODUCT({BJ},{AY},{AW},--({BL}=0))/SUMPRODUCT(--ISNUMBER({BJ}),{AY},{AW},--({BL}=0))", wavg(bj, m10)),
        ("Pa vlerat ekstreme", f"=SUMPRODUCT({BJ},{AY},{AW},--({BM}=0))/SUMPRODUCT(--ISNUMBER({BJ}),{AY},{AW},--({BM}=0))", wavg(bj, mx)),
        ("N - të gjitha intervistat", f"=SUMPRODUCT(--ISNUMBER({BJ}))", float(pd.to_numeric(d[bj], errors='coerce').notna().sum())),
        ("N - shpenzime ≤ 10,000 €", f"=SUMPRODUCT(--ISNUMBER({BJ}),--({BL}=0))", float((pd.to_numeric(d[bj], errors='coerce').notna() & m10).sum())),
        ("N - pa vlerat ekstreme", f"=SUMPRODUCT(--ISNUMBER({BJ}),--({BM}=0))", float((pd.to_numeric(d[bj], errors='coerce').notna() & mx).sum())),
        ("KRAHASIM (të gjitha intervistat, për person në ditë)", None, None),
        ("shmesatare (calculime)", f"=SUMPRODUCT({col_of('shmesatare')},{AY},{AW})/SUMPRODUCT(--ISNUMBER({col_of('shmesatare')}),{AY},{AW})", wavg("shmesatare")),
        ("shpz2 (calculime)", f"=SUMPRODUCT({col_of('shpz2')},{AY},{AW})/SUMPRODUCT(--ISNUMBER({col_of('shpz2')}),{AY},{AW})", wavg("shpz2")),
        ('Me shpenzimet para mbërritjes (si "Shpenzimi total" i raportit)', f"=SUMPRODUCT({col_of('Total shpenzimet për person/ditë')},{AY},{AW})/SUMPRODUCT(--ISNUMBER({col_of('Total shpenzimet për person/ditë')}),{AY},{AW})", wavg("Total shpenzimet për person/ditë")),
        ("TË DHËNA BAZË", None, None),
        ("Numri i intervistave", f"=COUNTA(A2:A{last})", float(n)),
        ("Numri i personave", f"=SUM({AW})", float(d["Nr. personave"].sum())),
        ("Mesatarja e ditëve të qëndrimit në Kosovë", f"=AVERAGE({AY})", float(d["Ditet e qendrimit ne KS"].mean())),
        ("Shpenzimet totale pa para mbërritjes (€)", f"=SUM({col_of('Shpenzimet pa para_arritjes')})", float(d["Shpenzimet pa para_arritjes"].sum())),
        ("Mesatarja e shpenzimeve për intervistë pa para mbërritjes (€)", f"=AVERAGE({col_of('Shpenzimet pa para_arritjes')})", float(d["Shpenzimet pa para_arritjes"].mean())),
        ("KUFIRI I VLERAVE EKSTREME (nga DB, përdoret edhe në A dhe M)", None, None),
        ("Q1", f"=QUARTILE({BJ},1)", q1),
        ("Q3", f"=QUARTILE({BJ},3)", q3),
        ("Kufiri i sipërm = Q3 + 3 × (Q3 - Q1)",
         f"={xl_col_to_name(tr)}21+3*({xl_col_to_name(tr)}21-{xl_col_to_name(tr)}20)" if db_threshold_ref is None else f"=DB!${BLOCK[0]}$22",
         q3 + 3 * (q3 - q1) if db_threshold_ref is None else db_threshold_ref),
    ]
    for k, (label, formula, value) in enumerate(block):
        ws.write_string(k + 1, tq, label, book.f["bold"] if formula is None else None)
        if formula:
            book.formula(ws, k + 1, tr, formula, value)
    ws.set_column(tq, tq, 55)
    return ws


# =============================================================================================  4. summary sheet
class Summary:
    """Writes the Permbledhje tables. Each cell = (Excel formula, value computed in Python)."""

    def __init__(self, book, db, A, M, cols, year, threshold, outliers, ext_limit):
        self.b, self.db, self.A, self.M = book, db, A, M
        self.ws = book.wb.add_worksheet("Permbledhje")
        self.L = letters(cols)
        self.last = len(db) + 1
        self.r = 0
        self.year, self.threshold, self.outliers, self.ext_limit = year, threshold, outliers, ext_limit
        self.ws.set_column(0, 0, 60)
        self.ws.set_column(1, 14, 14)

    # ---- ranges and Excel-like helpers
    def R(self, c, sheet="DB"):
        return f"{sheet}!${self.L[c]}$2:${self.L[c]}${self.last if sheet == 'DB' else {'A': len(self.A), 'M': len(self.M)}[sheet] + 1}"

    def put(self, c, cell, fmt=None):
        f, v = cell
        if isinstance(f, str) and f.startswith("="):
            self.b.formula(self.ws, self.r, c, f, v, fmt)
        elif isinstance(f, str):
            self.ws.write_string(self.r, c, f, self.b.f.get(fmt) if fmt else None)
        else:
            self.ws.write_number(self.r, c, float(f), self.b.f.get(fmt) if fmt else None)

    def text(self, s, fmt=None, col=0):
        self.ws.write_string(self.r, col, s, self.b.f.get(fmt) if fmt else None)

    def nl(self, k=1):
        self.r += k

    def cell(self, a1):
        return a1

    def here(self, col):
        return f"{xl_col_to_name(col)}{self.r + 1}"

    def T(self, title, *cols):
        """Title, marked when a column it needs is not in this year's export (the table then shows 0)."""
        miss = [c for c in cols if c in ABSENT]
        return title + (f" – KUJDES: s'ka të dhëna në eksportin {self.year} ({', '.join(miss)})" if miss else "")

    # ---- table types
    def freq(self, title, header, items, note=None, total_label="Total", blank_after_title=False):
        """items: [(label, formula, value)] → label | Frekuenca | Përqindja, with a Total row."""
        self.text(title, "tab"); self.nl()
        if note:
            self.text(note, "note"); self.nl()
        if blank_after_title:
            self.nl()
        for j, h in enumerate(header):
            self.text(h, "hdr", j)
        self.nl()
        first, tot_row = self.r + 1, self.r + len(items) + 1
        total = sum(v for _, _, v in items)
        for lab, f, v in items:
            self.text(str(lab), "lbl")
            self.put(1, (f, v), "int")
            self.put(2, (f"=IF($B${tot_row}=0,0,B{self.r + 1}/$B${tot_row})", v / total if total else 0), "pct")
            self.nl()
        self.text(total_label, "tot")
        self.put(1, (f"=SUM(B{first}:B{tot_row - 1})", total), "inttot")
        self.put(2, (f"=IF(B{tot_row}=0,0,B{tot_row}/B{tot_row})", 1 if total else 0), "pcttot")
        self.nl(2)
        return {lab: v for lab, _, v in items}, tot_row

    def stats(self, title, label, col, header_blank=False):
        s = pd.to_numeric(self.db[col], errors="coerce").dropna()
        self.text(title, "tab"); self.nl()
        if header_blank:
            self.nl()
        for j, h in zip([0, 1, 2, 3, 4, 7], ["Variabla", "Obs", "Mesatarja", "Dev.Std", "Min", "Max"]):
            self.text(h, "hdr", j)
        self.nl()
        R = self.R(col)
        self.text(label, "lbl")
        for j, (fn, v, fm) in zip([1, 2, 3, 4, 7], [("COUNT", len(s), "int"), ("AVERAGE", s.mean(), "dec"),
                                                   ("STDEV", s.std(ddof=1), "dec"), ("MIN", s.min(), "int"), ("MAX", s.max(), "int")]):
            self.put(j, (f"={fn}({R})", float(v)), fm)
        self.nl(2)

    # ---- data helpers
    def countif(self, col, val, df=None):
        df = self.db if df is None else df
        return f"=COUNTIF({self.R(col)},{val})", float((df[col] == val).sum())

    def bins(self, title, header, col, edges, top_label=None, lower_inclusive_zero=False, blank=False):
        x = pd.to_numeric(self.db[col], errors="coerce")
        items = []
        R = self.R(col)
        for k in range(len(edges) - 1):
            a, b = edges[k], edges[k + 1]
            lab = (f"{a:,}" if a == 0 else f"{a + 1:,}") + f"-{b:,}"
            if k == 0 and lower_inclusive_zero:
                items.append((lab, f'=COUNTIFS({R},">="&{a},{R},"<="&{b})', float(((x >= a) & (x <= b)).sum())))
            else:
                items.append((lab, f'=COUNTIFS({R},">"&{a},{R},"<="&{b})', float(((x > a) & (x <= b)).sum())))
        if top_label:
            items.append((top_label, f'=COUNTIF({R},">{edges[-1]}")', float((x > edges[-1]).sum())))
        return self.freq(title, header, items, blank_after_title=blank)

    # ---- the whole sheet
    def build(self):
        db, A, M, L, R = self.db, self.A, self.M, self.L, self.R
        N = len(db)
        self.text(f"Rezultatet e anketës {self.year} – shpenzimet e jo-rezidentëve në Kosovë", "title"); self.nl(3)
        self.text("Parametrat (përdoren edhe nga fletët DB, A, M)", "h1"); self.nl()
        self.text('Pragu për "shpenzime ≤ 10,000 €" (Total Shpenzimet, me para mbërritjes)')
        self.put(1, (self.threshold, None)); self.nl()
        self.text('Outlier-at e listës (referues; vetëm shënohen në kolonën "Outlier (lista)")')
        for j, o in enumerate(self.outliers[:12]):
            self.put(1 + j, (o, None))
        self.nl()
        self.text("Kufiri i vlerave ekstreme: Q3 + 3 × (Q3 − Q1) e shpenzimeve për person në ditë (DB)")
        self.put(1, (f"=DB!${BLOCK[0]}$22", self.ext_limit)); self.nl()
        self.text("Numri i intervistave (DB)")
        self.put(1, (f"=COUNTA(DB!$A$2:$A${self.last})", float(N))); self.nl()
        B8 = "$B$8"

        # ------------------------------------------------ 1. Struktura
        self.text("1. STRUKTURA E GRUPIT/INDIVIDIT TË UDHËTARËVE", "h1"); self.nl(2)
        self.freq(self.T("Tabela 1.1. Numri i të anketuarve sipas pikave kufitare", "pytesorivendi"), ["Pika kufitare", "Frekuenca", "Përqindja"],
                  [("Aeroport",) + self.countif("pytesorivendi", 2), ("Merdarë",) + self.countif("pytesorivendi", 1)]
                  + [(PIKAT[k][1],) + self.countif("pytesorivendi", k) for k in sorted(PIKAT) if k > 2])
        self.freq("Tabela 1.2 / Fig 1.1. Ndarja e të anketuarve sipas grupit", ["Grupi", "Frekuenca", "Përqindja"],
                  [("Individual",) + self.countif("group", 1), ("Grupor",) + self.countif("group", 2)])
        aw = db["Nr. personave"]
        items = [(str(k),) + self.countif("Nr. personave", k) for k in range(1, 11)]
        items.append(("Mbi 10", f'=COUNTIF({R("Nr. personave")},">10")', float((aw > 10).sum())))
        self.freq("Tabela 1.3. Numri i personave në grup", ["Nr personave", "Frekuenca", "Përqindja"], items,
                  note="Nr. personave = 1 për individët, nrpers për grupet (kolona AW në DB).")
        self.stats("Tabela 1.4. Mesatarja e madhësisë së grupit", "Nr. personave", "Nr. personave")
        G, I = db["group"], db["nrpers18"]
        adults = aw[G == 1].sum() + aw[(G == 2) & I.notna()].sum() - I[G == 2].sum()
        self.freq(self.T("Fig 1.2. Përqindja e personave mbi dhe nën moshën 18 vjeçare", "nrpers18"), ["Mosha", "Frekuenca", "Përqindja"],
                  [("Nr. personave mbi 18 vjet",
                    f'=(SUMIFS({R("Nr. personave")},{R("group")},1)+SUMPRODUCT(({R("group")}=2)*ISNUMBER({R("nrpers18")}),{R("Nr. personave")}))-SUMIFS({R("nrpers18")},{R("group")},2)',
                    float(adults)),
                   ("Nr. personave nën 18 vjet", f'=SUMIFS({R("nrpers18")},{R("group")},2)', float(I[G == 2].sum()))],
                  total_label="Total persona", blank_after_title=True)

        # countries: frequency desc, ties by first appearance
        J = db["shteti"]
        counts = J.dropna().astype(int).value_counts()
        first = {c: k for k, c in enumerate(dict.fromkeys(J.dropna().astype(int)))}
        codes = sorted(counts.index, key=lambda c: (-counts[c], first[c]))
        self.codes = codes
        lab = lambda c: f"{COUNTRIES[c][0]} – {COUNTRIES[c][1]}" if c in COUNTRIES else f"Kodi {c} (pa etiketë)"
        name = lambda c: COUNTRIES[c][1] if c in COUNTRIES else f"Kodi {c}"
        items = [(lab(c),) + self.countif("shteti", c) for c in codes]
        items.append(("Pa përgjigje", f"={B8}-COUNT({R('shteti')})", float(N - J.notna().sum())))
        self.freq(self.T("Tabela 1.5. Numri i të anketuarve sipas shteteve", "shteti"), ["Shteti", "Frekuenca", "Përqindja"], items,
                  blank_after_title=True)
        top2 = codes[:2]
        items = [(name(c),) + self.countif("shteti", c) for c in top2]
        items.append(("Shtetet tjera", f"={B8}-" + "-".join(f"COUNTIF({R('shteti')},{c})" for c in top2),
                      float(N - sum((J == c).sum() for c in top2))))
        self.freq(self.T("Fig 1.3. Përqindja e të anketuarve sipas shteteve kryesore", "shteti"), ["Shteti", "Frekuenca", "Përqindja"], items)
        K = db["pasaporta"]
        kos = aw[K == 1].sum() + db["miksPasaportKosovare"][K == 3].sum()
        huaj = aw[K == 2].sum() + db["miksPasaportaHuaj"][K == 3].sum()
        _, tot = self.freq(self.T("Tabela 1.6 / Fig 1.4. Numri i personave sipas llojit të pasaportave", "pasaporta", "miksPasaportKosovare", "miksPasaportaHuaj"), ["Pasaporta", "Frekuenca", "Përqindja"],
                           [("Kosovare", f'=SUMIFS({R("Nr. personave")},{R("pasaporta")},1)+SUMIFS({R("miksPasaportKosovare")},{R("pasaporta")},3)', float(kos)),
                            ("Të huaja", f'=SUMIFS({R("Nr. personave")},{R("pasaporta")},2)+SUMIFS({R("miksPasaportaHuaj")},{R("pasaporta")},3)', float(huaj))],
                           total_label="Total persona", blank_after_title=True)
        self.r -= 1
        self.put(1, (f"=SUM({R('Nr. personave')})", float(aw.sum()))); self.nl(3)
        self.freq(self.T("Tabela 1.7. Qëllimi i vizitës në Kosovë", "Qellimi"), ["Qëllimi", "Frekuenca", "Përqindja"],
                  [("Personale",) + self.countif("Qellimi", 1), ("Biznesore",) + self.countif("Qellimi", 2)])
        self.stats("Tabela 1.8. Mesatarja e ditëve të qëndrimit në Kosovë", "Ditët e qëndrimit në KS", "Ditet e qendrimit ne KS",
                   header_blank=True)
        self.freq("Fig 1.5. Vizitë jashtë Kosovës", ["Vizita", "Frekuenca", "Përqindja"],
                  [("Jo vizitë (vetëm në Kosovë)",) + self.countif("vizitejashtks", 1),
                   ("Vizitë jashtë Kosovës",) + self.countif("vizitejashtks", 2)])
        Q = db["vendi"]
        vcodes = sorted(Q.dropna().astype(int).unique())
        self.freq(self.T("Fig 1.6. Përqindja e vendeve më të vizituara", "vendi"), ["Vendi", "Frekuenca", "Përqindja"],
                  [(PLACES.get(c, f"Kodi {c} (pa etiketë)"),) + self.countif("vendi", c) for c in vcodes],
                  blank_after_title=True)

        # ------------------------------------------------ 2. Remitancat
        self.text("2. REMITANCAT NË KOSOVË", "h1"); self.nl(2)
        self.freq(self.T("Fig 2.1. Remitanca (a dërgoni para në Kosovë?)", "remitanca"), ["Përgjigja", "Frekuenca", "Përqindja"],
                  [("Po",) + self.countif("remitanca", 1), ("Jo",) + self.countif("remitanca", 2)])
        tl = ["Transfer bankar", "Kartel bankare", "Agjenci për transfer parash", "I sjell personalisht", "Tjera"]
        items = [(t, f"=SUM({R(f'transfer__{k + 1}')})", float(db[f"transfer__{k + 1}"].sum())) for k, t in enumerate(tl)]
        self.text(self.T("Fig 2.2. Format e dërgesave të remitancave", "transfer__1", "transfer__2", "transfer__3", "transfer__4", "transfer__5"), "tab"); self.nl(2)
        self.r -= 1
        _, t22 = self.freq_no_title(["", "Frekuenca", "Përqindja"], items, "Total përmendje")
        f0 = t22 - len(items)
        self.freq(self.T("Fig 2.3. Mënyrat e dërgimit të remitancave", "transfer__1", "transfer__2", "transfer__3", "transfer__4", "transfer__5"), ["Mënyra", "Frekuenca", "Përqindja"],
                  [("Formale", f"=SUM(B{f0}:B{f0 + 2})", sum(v for _, _, v in items[:3])),
                   ("Informale", f"=SUM(B{f0 + 3}:B{f0 + 4})", sum(v for _, _, v in items[3:]))],
                  total_label="Total përmendje", blank_after_title=True)

        # ------------------------------------------------ 3. Shpenzimet
        self.text("3. SHPENZIMET E JO-REZIDENTËVE NË KOSOVË", "h1"); self.nl(3)
        AZ, AI, AH = (db[c] for c in ["Total Shpenzimet", "blerjePatun", "ParaFalur"])
        self.money(self.T("Hyrje. Pjesa e kategorive kryesore në shpenzimet totale", "blerjaProd"), ["Kategoria", "Shuma (€)", "Përqindja"],
                   [("Shpenzimet e udhëtimit", f"=SUM({R('Total Shpenzimet')})-SUM({R('blerjePatun')})-SUM({R('ParaFalur')})", float(AZ.sum() - AI.sum() - AH.sum())),
                    ("Patundshmëri", f"=SUM({R('blerjePatun')})", float(AI.sum())),
                    ("Remitanca (para të falura)", f"=SUM({R('ParaFalur')})", float(AH.sum()))], blank_after_title=True)
        self.stats("Tabela 3.1. Mesatarja e shpenzimeve të udhëtimit (për intervistë)", "Total Shpenzimet (€)", "Total Shpenzimet")
        self.bins("Tabela 3.2. Shpenzimet e udhëtimit nën 10,000 €", ["Shpenzimet e udhëtimit nën 10,000 (€)", "Frekuenca", "Përqindja"],
                  "Total Shpenzimet", list(range(0, 10001, 500)), lower_inclusive_zero=True)
        self.r -= 1
        self.text("Mesatarja e shpenzimeve për intervistë, vetëm ≤ 10,000 € (€)")
        self.put(1, (f'=AVERAGEIFS({R("Total Shpenzimet")},{R("Total Shpenzimet")},"<="&$B$5)', float(AZ[AZ <= self.threshold].mean())), "dec")
        self.nl(2)
        self.per_day()
        self.bins("Tabela 3.3. Shpenzimet e udhëtimit mbi 10,000 €", ["Shpenzimet e udhëtimit mbi 10,000 (€)", "Frekuenca", "Përqindja"],
                  "Total Shpenzimet", [10000, 20000, 30000, 40000, 50000, 60000, 70000, 80000, 90000, 100000, 150000, 200000, 300000],
                  top_label="Mbi 300,000")
        self.money("Fig 3.1. Përqindja e shpenzimeve për veten ose grupin", ["Grupi", "Shpenzimet totale (€)", "Përqindja"],
                   [("Vetëm", f"=SUMIFS({R('Total Shpenzimet')},{R('group')},1)", float(AZ[G == 1].sum())),
                    ("Për grup", f"=SUMIFS({R('Total Shpenzimet')},{R('group')},2)", float(AZ[G == 2].sum()))])
        self.money(self.T("Fig 3.2. Shpenzimet e ndara sipas shteteve kryesore", "shteti"), ["Shteti", "Shpenzimet totale (€)", "Përqindja"],
                   [(name(c), f"=SUMIFS({R('Total Shpenzimet')},{R('shteti')},{c})", float(AZ[J == c].sum())) for c in top2]
                   + [("Shtetet tjera", f"=SUM({R('Total Shpenzimet')})-" + "-".join(f"SUMIFS({R('Total Shpenzimet')},{R('shteti')},{c})" for c in top2),
                       float(AZ.sum() - sum(AZ[J == c].sum() for c in top2)))])
        hb = [0, 100, 200, 300, 400, 500, 600, 700, 800, 900, 1000, 1500, 2000, 3000, 4000, 5000]
        self.bins("Tabela 3.4. Shpenzimet për shëndetësi", ["Shpenzimet për shëndetësi (€)", "Frekuenca", "Përqindja"],
                  "shpenzimeSh", hb, top_label="Mbi 5,000", blank=True)
        self.r -= 1
        self.text("Shuma totale e shpenzimeve për shëndetësi (€)")
        self.put(1, (f"=SUM({R('shpenzimeSh')})", float(db["shpenzimeSh"].sum())), "int"); self.nl(2)
        u = I > 0
        self.money(self.T("Fig 3.3. Shpenzimet kryesore të grupeve me persona nën 18 vjeç", "nrpers18", "blerjaProd"), ["Kategoria", "Shuma (€)", "Përqindja"],
                   [(lab_, f'=SUMIFS({R(c)},{R("nrpers18")},">0")', float(db[c][u].sum())) for lab_, c in
                    [("Blerje të produkteve", "blerjaProd"), ("Ushqim", "Ushqim"), ("Argëtim", "argetimi"), ("Transport", "shpenzimetTran")]],
                   blank_after_title=True)
        self.r -= 1
        self.text("Numri i personave nën 18 vjeç"); self.put(1, (f"=SUM({R('nrpers18')})", float(I.sum())), "int"); self.nl()
        self.text("Numri i intervistave me persona nën 18 vjeç"); self.put(1, (f'=COUNTIF({R("nrpers18")},">0")', float(u.sum())), "int"); self.nl(2)
        pb = [0, 100, 200, 300, 400, 500, 600, 700, 800, 900, 1000, 1500, 2000, 3000, 4000, 5000, 6000, 10000]
        self.bins("Tabela 3.5. Para të falura nën 10,000 €", ["Para të falura nën 10,000 (€)", "Frekuenca", "Përqindja"], "ParaFalur", pb, blank=True)
        self.r -= 1
        self.text("Intervista me para të falura mbi 10,000 €"); self.put(1, (f'=COUNTIF({R("ParaFalur")},">10000")', float((AH > 10000).sum())), "int"); self.nl(2)

        # ------------------------------------------------ 4–5
        AM = pd.to_numeric(db["dhuratangaKosova"], errors="coerce")
        self.text(self.T("4. PRANIMET NGA KOSOVA (dhurata nga kosovarët)", "dhuratangaKosova"), "h1"); self.nl()
        self.text("Treguesi", "hdr"); self.text("Vlera", "hdr", 1); self.nl()
        for lab_, f, v, fm in [
            ("Intervista që kanë pranuar dhurata (> 0)", f'=COUNTIF({R("dhuratangaKosova")},">0")', float((AM > 0).sum()), "int"),
            ("% e të gjitha intervistave", f'=COUNTIF({R("dhuratangaKosova")},">0")/{B8}', (AM > 0).sum() / N, "pct"),
            ("% e intervistave që iu përgjigjën pyetjes", f'=COUNTIF({R("dhuratangaKosova")},">0")/COUNT({R("dhuratangaKosova")})', (AM > 0).sum() / AM.notna().sum(), "pct"),
            ("Vlera totale e dhuratave të pranuara (€)", f"=SUM({R('dhuratangaKosova')})", float(AM.sum()), "int"),
            ("Mesatarja për ata që kanë pranuar (€)", f'=AVERAGEIF({R("dhuratangaKosova")},">0")', float(AM[AM > 0].mean()), "dec")]:
            self.text(lab_, "lbl"); self.put(1, (f, v), fm); self.nl()
        self.nl()
        self.text("5. DEPOZITAT NË KOSOVË", "h1"); self.nl(2)
        self.freq(self.T("Fig 5.1. Depozitat në Kosovë (a keni llogari bankare në Kosovë?)", "deposita"), ["Përgjigja", "Frekuenca", "Përqindja"],
                  [("Po",) + self.countif("deposita", 1), ("Jo",) + self.countif("deposita", 2)])
        self.freq(self.T("Fig 5.2. Format e depozitave", "forma"), ["Forma", "Frekuenca", "Përqindja"],
                  [("Cash",) + self.countif("forma", 1), ("Përmes bankës",) + self.countif("forma", 2), ("Të dyja",) + self.countif("forma", 3)],
                  blank_after_title=True)

        # ------------------------------------------------ Aneks
        self.text("ANEKS", "h1"); self.nl(2)
        S = db["remitanca"]
        top3 = sorted(codes[:3])
        self.freq(self.T("Aneks 1 / Fig 1. Dërgesat e remitancave sipas shteteve (intervistat që dërgojnë para)", "remitanca", "shteti"), ["Shteti", "Frekuenca", "Përqindja"],
                  [(name(c), f"=COUNTIFS({R('remitanca')},1,{R('shteti')},{c})", float(((S == 1) & (J == c)).sum())) for c in top3]
                  + [("Tjerat", f"=COUNTIF({R('remitanca')},1)-" + "-".join(f"COUNTIFS({R('remitanca')},1,{R('shteti')},{c})" for c in top3),
                      float((S == 1).sum() - sum(((S == 1) & (J == c)).sum() for c in top3)))])
        cats = [("Shpenzimet e udhëtimit para se të arrijnë në Kosovë", "para_arritjes"), ("Blerje e produkteve", "blerjaProd"),
                ("Akomodimi", "akomodimi"), ("Ushqim", "Ushqim"), ("Argëtim", "argetimi"),
                ("Shpenzimet e transportit në Kosovë", "shpenzimetTran"), ("Shpenzime për shëndetësi", "shpenzimeSh"),
                ("Dhuratat për kosovarët", "DhurataPer"), ("Para që ju keni falur familjarëve", "ParaFalur"),
                ("Blerje e patundshmërisë", "blerjePatun"), ("Shpenzime të tjera", "vleraShpenzimit")]
        self.money(self.T("Aneks 2 / Tabela 1. Shpenzimet dhe ndikimi i kategorive mbi shpenzimet totale", "blerjaProd"), ["Lloji i shpenzimeve", "Shuma në euro", "Përqindja"],
                   [(l_, f"=SUM({R(c)})", float(db[c].sum())) for l_, c in cats])
        self.annex3(codes, lab)
        return self

    def freq_no_title(self, header, items, total_label):
        for j, h in enumerate(header):
            if h:
                self.text(h, "hdr", j)
        self.nl()
        first, tot_row = self.r + 1, self.r + len(items) + 1
        total = sum(v for _, _, v in items)
        for lab, f, v in items:
            self.text(lab, "lbl"); self.put(1, (f, v), "int")
            self.put(2, (f"=IF($B${tot_row}=0,0,B{self.r + 1}/$B${tot_row})", v / total if total else 0), "pct"); self.nl()
        self.text(total_label, "tot")
        self.put(1, (f"=SUM(B{first}:B{tot_row - 1})", total), "inttot")
        self.put(2, (f"=IF(B{tot_row}=0,0,B{tot_row}/B{tot_row})", 1 if total else 0), "pcttot")
        self.nl(2)
        return None, tot_row

    def money(self, title, header, items, blank_after_title=False):
        self.text(title, "tab"); self.nl()
        if blank_after_title:
            self.nl()
        for j, h in enumerate(header):
            self.text(h, "hdr", j)
        self.nl()
        first, tot_row = self.r + 1, self.r + len(items) + 1
        total = sum(v for _, _, v in items)
        for lab, f, v in items:
            self.text(lab, "lbl"); self.put(1, (f, v), "int")
            self.put(2, (f"=IF($B${tot_row}=0,0,B{self.r + 1}/$B${tot_row})", v / total if total else 0), "pct"); self.nl()
        self.text("Total", "tot")
        self.put(1, (f"=SUM(B{first}:B{tot_row - 1})", total), "inttot")
        self.put(2, (f"=IF(B{tot_row}=0,0,1)", 1 if total else 0), "pcttot")
        self.nl(2)

    def per_day(self):
        """Spending per person per day block (rows 182–196 in 2026)."""
        db, A, M = self.db, self.A, self.M
        self.text("Shpenzimet për person në ditë", "tab"); self.nl(2)
        hdr = ["Treguesi (€ për person në ditë)", "DB (Total)", "A (Aeroport)", "M (Merdarë)"]
        for j, h in enumerate(hdr):
            self.text(h, "hdr", j)
        self.text("Mesatarja per raport (pa-para mberritjes)", "hdr", 7)
        self.text("Mesatarja per raport (pa-para-mberritjes, dhe me vlera ekstreme te hequra", "hdr", 9)
        self.nl()
        blocks = {s: self.blockvals[s] for s in ("DB", "A", "M")}
        rows = [("Pa para mbërritjes – të gjitha intervistat", 3), ("Pa para mbërritjes – vetëm shpenzime ≤ 10,000 €", 4),
                ("Pa para mbërritjes – pa vlerat ekstreme", 5), ("shmesatare (calculime)", 10), ("shpz2 (calculime)", 11),
                ("Me shpenzimet para mbërritjes", 12)]
        row_of = {}
        for lab, k in rows:
            self.text(lab)
            for j, s in enumerate(("DB", "A", "M")):
                self.put(1 + j, (f"={s}!{BLOCK[0]}{k}", blocks[s][k]), "dec")
            row_of[k] = self.r + 1
            if k == 3:
                self.put(7, (f"=ROUND(B{self.r + 1},1)", round(blocks['DB'][3], 1)), "dec")
            if k == 5:
                self.put(9, (f"=ROUND(B{self.r + 1},1)", round(blocks['DB'][5], 1)), "dec")
            if k == 10:
                for j, h in zip([5, 6, 7, 9], ["Airport", "Merdare", "Mesatarja si 2024", "Sa anketa perfshihen"]):
                    self.text(h, "hdr", j)
            if k == 11:
                for j, (pt, df) in zip([5, 6, 7], [("A", A), ("M", M), (None, db)]):
                    f, v = self.si2024(pt)
                    self.put(j, (f, v), "dec")
                f, v = self.si2024(None, count=True)
                self.put(9, (f, v), "int")
            self.nl()
        # pa para mbërritjes dhe pa patundshmëri – ≤ 10,000
        self.text("Pa para mbërritjes dhe pa patundshmëri – vetëm shpenzime ≤ 10,000 €")
        Rr = self.R
        BJ, BL, BI, AI, AY, AW = (Rr(c) for c in ["Shpenzimet pa para_arritjes për person/ditë", "Total Shpenzimet > 10000",
                                                  "Shpenzimet pa para_arritjes", "blerjePatun", "Ditet e qendrimit ne KS", "Nr. personave"])
        m = pd.to_numeric(db["Shpenzimet pa para_arritjes për person/ditë"], errors="coerce").notna() & (db["Total Shpenzimet > 10000"] == 0)
        v = (db["Shpenzimet pa para_arritjes"][m] - db["blerjePatun"].fillna(0)[m]).sum() / (db["Ditet e qendrimit ne KS"][m] * db["Nr. personave"][m]).sum()
        self.put(1, (f"=SUMPRODUCT(--ISNUMBER({BJ}),--({BL}=0),{BI}-{AI})/SUMPRODUCT(--ISNUMBER({BJ}),--({BL}=0),{AY},{AW})", v), "dec")
        self.nl()
        for j, h in enumerate(["Numri i intervistave në llogaritje", "DB (Total)", "A (Aeroport)", "M (Merdarë)"]):
            self.text(h, "hdr", j)
        self.nl()
        for lab, k in [("Të gjitha intervistat", 6), ("Shpenzime ≤ 10,000 €", 7), ("Pa vlerat ekstreme", 8)]:
            self.text(lab)
            for j, s in enumerate(("DB", "A", "M")):
                self.put(1 + j, (f"={s}!{BLOCK[0]}{k}", blocks[s][k]), "int")
            self.nl()
        f, v = self.si2024(None, aggregate=True)
        self.put(8, (f, v), "dec")
        self.nl()

    def si2024(self, point, count=False, aggregate=False):
        """'Mesatarja si 2024': mean over valid interviews of in-Kosovo spending / (persons × days in Kosovo)."""
        db, L, last = self.db, self.L, self.last
        rg = lambda c: f"DB!${L[c]}$2:${L[c]}${last}"
        days = f"({rg('Diferenca ne data')}-({rg('vizitejashtks')}=2)*{rg('ditet')})"
        cond = (f"({rg('interview__key')}<>\"\")" + (f"*({rg('Pika kufitare')}=\"{point}\")" if point else "")
                + f"*({rg('totali')}<>\"\")*({rg('totali')}<=$B$5)*({rg('vizitejashtks')}<>\"\")"
                + f"*((({rg('vizitejashtks')}<>2)+({rg('ditet')}<>\"\"))>0)*({days}>0)")
        spend = "(" + "+".join(rg(c) for c in IN_KOSOVO) + ")"
        df = db if point is None else db[db["Pika kufitare"] == point]
        num = lambda c: pd.to_numeric(df[c], errors="coerce").fillna(0)
        d = df["Diferenca ne data"] - (df["vizitejashtks"] == 2) * num("ditet")
        ok = (df["totali"].notna() & (df["totali"] <= self.threshold) & df["vizitejashtks"].notna()
              & ((df["vizitejashtks"] != 2) | df["ditet"].notna()) & (d > 0))
        sp = sum(num(c) for c in IN_KOSOVO)
        if count:
            return f"=SUMPRODUCT({cond})", float(ok.sum())
        if aggregate:
            return (f"=SUMPRODUCT({cond}*{spend})/SUMPRODUCT({cond}*{rg('Nr. personave')}*{days})",
                    float(sp[ok].sum() / (df["Nr. personave"][ok] * d[ok]).sum()))
        return (f"=SUMPRODUCT({cond}*{spend}/({rg('Nr. personave')}*{days}+1-{cond}))/SUMPRODUCT({cond})",
                float((sp / (df["Nr. personave"] * d))[ok].mean()))

    def annex3(self, codes, lab):
        db, R = self.db, self.R
        self.text("Aneks 3 / Tabela 2. Shpenzimet totale sipas shteteve dhe kategorive (€)", "tab"); self.nl()
        heads = [("Para mbërritjes", "para_arritjes"), ("Produkte", "blerjaProd"), ("Akomodimi", "akomodimi"), ("Ushqim", "Ushqim"),
                 ("Argëtimi", "argetimi"), ("Transport", "shpenzimetTran"), ("Shëndetësi", "shpenzimeSh"), ("Dhurata", "DhurataPer"),
                 ("Para të falura", "ParaFalur"), ("Patundshmëri", "blerjePatun"), ("Tjera", "vleraShpenzimit"),
                 ("Shpenzimet totale", "Total Shpenzimet")]
        cols = [1, 2, 3, 4, 7, 8, 9, 10, 11, 12, 13, 14]
        self.text("Shteti", "hdr")
        for j, (h, _) in zip(cols, heads):
            self.text(h, "hdr", j)
        self.nl()
        first = self.r + 1
        J = db["shteti"]
        for c in codes + [None]:
            self.text(lab(c) if c is not None else "Pa përgjigje", "lbl")
            for j, (_, col) in zip(cols, heads):
                crit = str(c) if c is not None else '""'
                m = (J == c) if c is not None else J.isna()
                self.put(j, (f"=SUMIFS({R(col)},{R('shteti')},{crit})", float(db[col][m].sum())), "int")
            self.nl()
        self.text("Total", "tot")
        for j, (_, col) in zip(cols, heads):
            cl = xl_col_to_name(j)
            self.put(j, (f"=SUM({cl}{first}:{cl}{self.r})", float(db[col].sum())), "inttot")
        self.nl()


def block_values(df, ext_limit):
    """Python values of the indicator block (BR3…BR22) for one sheet, by row number."""
    bj = pd.to_numeric(df["Shpenzimet pa para_arritjes për person/ditë"], errors="coerce")

    def w(col, mask=None):
        v = pd.to_numeric(df[col], errors="coerce")
        m = v.notna() if mask is None else v.notna() & mask
        ww = df["Ditet e qendrimit ne KS"][m] * df["Nr. personave"][m]
        return float((v[m] * ww).sum() / ww.sum()) if ww.sum() else np.nan

    m10, mx = df["Total Shpenzimet > 10000"] == 0, df["Vlerë ekstreme (shpenzimet pa para_arritjes)"] == 0
    return {3: w("Shpenzimet pa para_arritjes për person/ditë"), 4: w("Shpenzimet pa para_arritjes për person/ditë", m10),
            5: w("Shpenzimet pa para_arritjes për person/ditë", mx), 6: float(bj.notna().sum()),
            7: float((bj.notna() & m10).sum()), 8: float((bj.notna() & mx).sum()), 10: w("shmesatare"), 11: w("shpz2"),
            12: w("Total shpenzimet për person/ditë")}


# =============================================================================================  5. enumerators
def write_enumerators(book, db, cols):
    ws = book.wb.add_worksheet("Anketuesit")
    L, last = letters(cols), len(db) + 1
    R = lambda c: f"DB!${L[c]}$2:${L[c]}${last}"
    ws.write_string(0, 0, "Efekti i anketuesit - krahasim nga fleta DB", book.f["title"])
    hdr = ["Pika kufitare", "anketuesi", "Nr. intervistash", "Treguesi kryesor - mes. e ponderuar",
           "Mes. e ponderuar e pikës pa këtë anketues", "% Diferenca = 0", "% Diferenca > 0", "% Diferenca < 0",
           '% intervista me ≥1 "nuk di"', 'Mes. "nuk di" për intervistë', "Numër në shpenzimeTjera (fushë teksti)",
           "Numër ≠ 0 në tetjera (fushë teksti)", "shumaShpenzimit = totali"]
    for j, h in enumerate(hdr):
        ws.write_string(2, j, h, book.f["hdr"])
    ws.set_column(0, 12, 16)
    F, E = db["anketuesi"], db["pytesorivendi"]
    BJ = pd.to_numeric(db["Shpenzimet pa para_arritjes për person/ditë"], errors="coerce")
    WT = db["Ditet e qendrimit ne KS"] * db["Nr. personave"]
    BA, BO = pd.to_numeric(db["Diferenca"], errors="coerce"), db['Nr. përgjigjesh "nuk di" (Final)']
    isnum = lambda s: s.apply(is_num)
    boR = R('Nr. përgjigjesh "nuk di" (Final)')
    r = 3
    enums = sorted(F.dropna().astype(int).unique())
    for point, code in [("A", 2), ("M", 1), ("Të dyja", None)]:
        for e in enums:
            sel = (F == e) & ((E == code) if code else True)
            other = ((E == code) if code else (F > 0)) & (F != e)
            selF = f"({R('anketuesi')}={e})*" + (f"({R('pytesorivendi')}={code})" if code else "1")   # "*1" so SUMPRODUCT counts it
            othF = (f"({R('pytesorivendi')}={code})" if code else f"({R('anketuesi')}>0)") + f"*({R('anketuesi')}<>{e})"
            bjR, ayR, awR = R("Shpenzimet pa para_arritjes për person/ditë"), R("Ditet e qendrimit ne KS"), R("Nr. personave")
            n = float(sel.sum())
            wav = lambda m: float((BJ[m & BJ.notna()] * WT[m & BJ.notna()]).sum() / WT[m & BJ.notna()].sum()) if WT[m & BJ.notna()].sum() else np.nan
            vals = [
                (f"=SUMPRODUCT({selF})", n),
                (f"=SUMPRODUCT({selF}*ISNUMBER({bjR}),{bjR},{ayR},{awR})/SUMPRODUCT({selF}*ISNUMBER({bjR}),{ayR},{awR})", wav(sel)),
                (f"=SUMPRODUCT({othF}*ISNUMBER({bjR}),{bjR},{ayR},{awR})/SUMPRODUCT({othF}*ISNUMBER({bjR}),{ayR},{awR})", wav(other)),
            ]
            for op, test in [("=0", BA == 0), (">0", BA > 0), ("<0", BA < 0)]:
                vals.append((f"=SUMPRODUCT({selF}*ISNUMBER({R('Diferenca')})*({R('Diferenca')}{op}))/SUMPRODUCT({selF})",
                             float((sel & BA.notna() & test).sum() / n) if n else np.nan))
            vals.append((f"=SUMPRODUCT({selF}*({boR}>0))/SUMPRODUCT({selF})",
                         float((sel & (BO > 0)).sum() / n) if n else np.nan))
            vals.append((f"=SUMPRODUCT({selF},{boR})/SUMPRODUCT({selF})",
                         float(BO[sel].sum() / n) if n else np.nan))
            vals.append((f"=SUMPRODUCT({selF}*ISNUMBER({R('shpenzimeTjera')}))", float((sel & isnum(db["shpenzimeTjera"])).sum())))
            tet = db["tetjera"]
            vals.append((f"=SUMPRODUCT({selF}*ISNUMBER({R('tetjera')})*({R('tetjera')}<>0))",
                         float((sel & isnum(tet) & (pd.to_numeric(tet.where(isnum(tet)), errors="coerce") != 0)).sum())))
            AO, Y = pd.to_numeric(db["shumaShpenzimit"], errors="coerce"), db["totali"]
            vals.append((f"=SUMPRODUCT({selF}*ISNUMBER({R('shumaShpenzimit')})*({R('shumaShpenzimit')}={R('totali')}))",
                         float((sel & AO.notna() & (AO == Y)).sum())))
            ws.write_string(r, 0, point)
            ws.write_number(r, 1, e)
            for j, (f, v) in enumerate(vals):
                book.formula(ws, r, 2 + j, f, v, "pct" if 5 <= 2 + j <= 8 else "dec")
            r += 1


# =============================================================================================  main
def main():
    ap = argparse.ArgumentParser(description="Clean the raw survey export into the final analysis workbook.")
    ap.add_argument("raw", help="raw export (.xlsx), first sheet")
    ap.add_argument("--dalja", help="output workbook (default: FINALEDataset_<year>_me_rezultatet_e_raportit.xlsx)")
    ap.add_argument("--viti", type=int, help="survey year (default: from interview dates)")
    ap.add_argument("--pragu", type=float, help="threshold for 'shpenzime ≤ 10,000 €' (default: skema.json)")
    ap.add_argument("--outliers", nargs="*", type=int, default=[], help="'referues' numbers to flag as outliers (list)")
    ap.add_argument("--korrigjime", help="CSV of manual corrections: interview__key,kolona,vlera,arsyeja")
    ap.add_argument("--codebook", help="codebook .xlsx – the sheet of the survey year (or the latest earlier one) is used "
                                       "for border-point, country and place codes")
    ap.add_argument("--skema", default=os.path.join(os.path.dirname(os.path.abspath(__file__)), "skema.json"),
                    help="year rules: column names, codes, labels (default: skema.json next to this script)")
    ap.add_argument("--fleta", default=0, help="sheet of the raw export (name or number, default: the first)")
    ap.add_argument("--pika", choices=["A", "M"], help="only if the export has no border-point column: A or M for all")
    a = ap.parse_args()
    r = pastro(**vars(a))
    print(f"Viti {r['viti']}: {r['n_raw']} intervista → {r['n_db']} në DB ({r['n_A']} A, {r['n_M']} M), "
          f"{r['n_ex']} të përjashtuara")
    print("Ndryshimet:\n" + "\n".join(f"  {v:5d}  {k}" for k, v in r["ndryshimet"].items()))
    if r["paralajmerime"]:
        print(f"\n{len(r['paralajmerime'])} paralajmërime për kontroll (shih fletën Log):")
        for w in r["paralajmerime"]:
            print("  ", w)
    print(f"\nRuajtur: {r['dalja']}")


def pastro(raw, dalja=None, viti=None, pragu=None, outliers=(), korrigjime=None, codebook=None, skema=None, fleta=0,
           pika=None):
    """Clean one raw export and write the final workbook. Returns a summary dict (also used by bqk.py and app.py)."""
    from types import SimpleNamespace
    a = SimpleNamespace(raw=raw, dalja=dalja, viti=viti, pragu=pragu, outliers=list(outliers or []),
                        korrigjime=korrigjime, codebook=codebook, fleta=fleta, pika=pika,
                        skema=skema or os.path.join(os.path.dirname(os.path.abspath(__file__)), "skema.json"))
    ABSENT.clear()
    if not os.path.exists(a.skema):
        raise SystemExit(f"Nuk u gjet {a.skema} – vendoseni skema.json në të njëjtin folder me skriptin ose jepni --skema")
    load_schema(a.skema)
    a.pragu = a.pragu if a.pragu is not None else DEFAULT_THRESHOLD
    log = Log()
    fleta = int(a.fleta) if str(a.fleta).isdigit() else a.fleta
    raw = pd.read_excel(a.raw, sheet_name=fleta) if not a.raw.lower().endswith(".csv") else pd.read_csv(a.raw)
    raw = raw.dropna(how="all").reset_index(drop=True)
    raw = harmonise_columns(raw, log)
    if "pytesorivendi" in ABSENT and a.pika:          # export of one border point only, without the column
        raw["pytesorivendi"] = a.pika
        ABSENT.discard("pytesorivendi")
        log.add("—", "pytesorivendi", "", a.pika, "Eksporti nuk ka kolonën e pikës – e gjithë skedari u shënua si "
                + {"A": "Aeroport", "M": "Merdarë"}.get(a.pika, a.pika) + " (zgjedhur nga përdoruesi)")
    elif "pytesorivendi" in ABSENT:
        log.add("—", "pytesorivendi", "", "", "KUJDES: eksporti nuk ka kolonën e pikës kufitare – tabelat dhe grafikët "
                "sipas Aeroport/Merdarë dalin bosh. Nëse skedari është vetëm për një pikë, zgjidhni pikën (--pika A ose M)")
    year = a.viti or guess_year(raw)
    pika_num, extra_names = set_year_codes(year, a.codebook, log)
    raw["_nukdi"] = raw.isin(MISSING_CODES).sum(axis=1)
    raw = harmonise_values(raw, year, pika_num, extra_names, log)
    nukdi = raw.pop("_nukdi")
    df = clean(raw.assign(_nukdi=nukdi), log, a.korrigjime)
    der = derive(df, a.outliers, a.pragu)
    full = pd.concat([df.drop(columns=["_nukdi"]), der], axis=1)
    full["_reason"] = full.apply(exclusion_reason, axis=1)
    db = full[full["_reason"].isna()].drop(columns="_reason").reset_index(drop=True)
    ex = full[full["_reason"].notna()].rename(columns={"_reason": "Arsyeja e përjashtimit"}).reset_index(drop=True)
    for _, rr in ex.iterrows():
        log.add(rr["interview__key"], "—", "", "", "Përjashtuar: " + rr["Arsyeja e përjashtimit"])
    # extreme-value limit (from DB) → flag on every sheet
    bj = "Shpenzimet pa para_arritjes për person/ditë"
    q1, q3 = quartile_inc(db[bj], 0.25), quartile_inc(db[bj], 0.75)
    ext = q3 + 3 * (q3 - q1)
    flag = lambda d: (pd.to_numeric(d[bj], errors="coerce") > ext).astype(int)
    db["Vlerë ekstreme (shpenzimet pa para_arritjes)"] = flag(db)
    ex["Vlerë ekstreme (shpenzimet pa para_arritjes)"] = flag(ex)
    cols = list(raw.columns) + DERIVED
    BLOCK[0] = xl_col_to_name(len(cols) + 2)
    db, ex = db[cols], ex[cols + ["Arsyeja e përjashtimit"]]
    A = db[db["Pika kufitare"] == "A"].reset_index(drop=True)
    M = db[db["Pika kufitare"] == "M"].reset_index(drop=True)
    others = {PIKAT[k][0]: db[db["Pika kufitare"] == PIKAT[k][0]].reset_index(drop=True) for k in sorted(PIKAT) if k > 2}
    sanity_checks(db, log)
    out = a.dalja or f"FINALEDataset_{year}_me_rezultatet_e_raportit.xlsx"

    book = Book(out)
    summ = Summary(book, db, A, M, cols, year, a.pragu, a.outliers, ext)
    summ.blockvals = {"DB": block_values(db, ext), "A": block_values(A, ext), "M": block_values(M, ext)}
    summ.build()
    write_enumerators(book, db, cols)
    write_data_sheet(book, "DB", db, cols)
    write_data_sheet(book, "A", A, cols, db_threshold_ref=ext)
    write_data_sheet(book, "M", M, cols, db_threshold_ref=ext)
    for L_, d_ in others.items():
        write_data_sheet(book, L_, d_, cols, db_threshold_ref=ext)
    write_data_sheet(book, "Perjashtuar", ex, cols, extra_col="Arsyeja e përjashtimit")
    ws = book.wb.add_worksheet("Log")
    for j, h in enumerate(["interview__key", "kolona", "vlera e vjetër", "vlera e re", "rregulli"]):
        ws.write_string(0, j, h, book.f["hdr"])
    for i, row in enumerate(log.rows):
        for j, v in enumerate(row):
            ws.write_string(i + 1, j, v)
    ws.set_column(0, 3, 16); ws.set_column(4, 4, 90)
    book.wb.close()

    rules = pd.Series([r[4].split(":")[0].split(" (")[0].split(" – ")[0] for r in log.rows]).value_counts()
    return {"viti": year, "dalja": out, "n_raw": len(raw), "n_db": len(db), "n_A": len(A), "n_M": len(M), "n_ex": len(ex),
            "pikat_tjera": {k: len(v) for k, v in others.items()}, "ndryshimet": rules.to_dict(),
            "paralajmerime": [r for r in log.rows if r[4].startswith("KUJDES")],
            "perjashtuar": [(r["interview__key"], r["Arsyeja e përjashtimit"]) for _, r in ex.iterrows()]}


if __name__ == "__main__":
    main()
