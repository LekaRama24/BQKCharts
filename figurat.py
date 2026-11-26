"""
figurat.py — the list of figures for the non-resident spending report, drawn with bqk_grafiket.py.

Run:
    python figurat.py Dataset_2027.xlsx                      # all figures → ./figurat_2027/
    python figurat.py Dataset_2027.xlsx --dalja dosja        # choose the output folder
    python figurat.py Dataset_2027.xlsx --vetem fig_3 fig_6  # only some figures (file name or title prefix)
    python figurat.py Dataset_2027.xlsx --viti 2027 --periudha "korrik–gusht"

The year is read from the interview dates in the DB sheet unless --viti is given.

Each figure is one dict in FIGURAT:
    file      output file name
    type      donut | hbar | histogram | vbar | share_bar | share_rows | grouped_hbar   (see bqk_grafiket.py)
    title     title printed on the chart
    table     title of the table in the "Permbledhje" sheet (start of the text is enough), OR
    calc      a function(ds) returning the data, for figures with no summary table (sections 4, 6, 7)
    subtitle  text; "{n}" is replaced by the sum of the values. Can also be a function(ds, data).
  optional:
    col=2                 read the 2nd value column (shares) instead of the 1st
    labels=[...]          replace the sheet labels, in order
    rename={old: new}     rename some labels
    sort="desc"           largest first;  keep_last="Tjerat" keeps that label at the bottom
    trim=True             drop empty intervals at the end of a histogram
    scale=100             multiply values (shares → percent)
    colors=[...]          palette names or hex, slice/bar order
    color="teal"          single colour (hbar / histogram)
    value_format          count | pct_count | pct | share | euro | euro1   (bar labels)
    figsize=(w, h), left=0.3, highlight_max=False, bold_values=True, ylabel="..."
"""
import argparse
import pandas as pd
import bqk_grafiket as bg
from bqk_grafiket import fmt

# Names for codes without labels in the summary sheet (from the codebook).
RELABEL = {
    "Kroaci (kodi 3, për t'u konfirmuar)": "Kroaci",
    "Kodi 6 (pa etiketë)": "Greqi",
    "Kodi 28 (pa etiketë)": "TR – Turqi",
}


# =============================================================================================  calculations
def under_header(ds, header):
    """Value directly below the cell `header` in the summary sheet (e.g. 'Mesatarja si 2024')."""
    for i, r in enumerate(ds.rows):
        for j, v in enumerate(r):
            if isinstance(v, str) and v.strip() == header:
                return ds.rows[i + 1][j]
    raise KeyError(header)


def gifts(ds):
    """4.1: received gifts yes/no, base = those who answered (Po / share of respondents)."""
    po = ds.value("Intervista që kanë pranuar dhurata")
    base = round(po / ds.value("% e intervistave që iu përgjigjën pyetjes"))
    return [("Po", po), ("Jo", base - po)]


def card(ds):
    """6.1: respondents who paid by card at least once vs cash only (DB sheet)."""
    cv = pd.to_numeric(ds.sheet("DB")["PaguarKartel"], errors="coerce")
    return [("Ka paguar me kartelë", int((cv > 0).sum())), ("Vetëm para të gatshme", int((cv == 0).sum()))]


def card_subtitle(ds, data):
    db = ds.sheet("DB")
    cv = pd.to_numeric(db["PaguarKartel"], errors="coerce")
    ok = cv.notna()
    share = 100 * cv[ok].sum() / db.loc[ok, "Total Shpenzimet"].sum()
    return (f"n = {fmt(ok.sum())} të anketuar   ·   pagesat me kartelë: {share:.1f}% e vlerës totale "
            "të shpenzimeve të tyre")


def ppd(ds):
    """6.2: spending per person per day, as in the summary sheet ('Mesatarja si 2024')."""
    return [("Gjithsej", under_header(ds, "Mesatarja si 2024")),
            ("Aeroport", under_header(ds, "Airport")),
            ("Merdarë", under_header(ds, "Merdare"))]


def ppd_one(df):
    """Same method for any subset of interviews: mean of (in-Kosovo spending) / (persons × days in Kosovo),
    interviews with reported total ≤ 10,000 € and days > 0."""
    num = lambda c: pd.to_numeric(df[c], errors="coerce").fillna(0)
    days = df["Diferenca ne data"] - (df["vizitejashtks"] == 2) * num("ditet")
    ok = (df["totali"].notna() & (df["totali"] <= 10000) & df["vizitejashtks"].notna()
          & ((df["vizitejashtks"] != 2) | df["ditet"].notna()) & (days > 0))
    spend = sum(num(c) for c in ["akomodimi", "blerjaProd", "Ushqim", "argetimi", "shpenzimetTran",
                                 "shpenzimeSh", "vleraShpenzimit"])
    return (spend / (df["Nr. personave"] * days))[ok].mean(), int(ok.sum())


COUNTRY_CODES = {2: "Gjermani", 1: "Zvicër", 6: "Austri", 20: "Slloveni", 9: "Belgjikë", 15: "Itali", 4: "Suedi",
                 11: "Francë", 5: "Britani e Madhe", 3: "SHBA", 8: "Norvegji", 12: "Holandë"}
MIN_INTERVIEWS = 10


def ppd_by_country(ds):
    """6.3: per person per day by country, countries with ≥ MIN_INTERVIEWS valid interviews."""
    db = ds.sheet("DB")
    rows = []
    for code, name in COUNTRY_CODES.items():
        m, n = ppd_one(db[db["shteti"] == code])
        if n >= MIN_INTERVIEWS:
            rows.append((name, m, n))
    return sorted(rows, key=lambda r: -r[1])


def totals_by_border(ds):
    return [("Aeroport", ds.sheet("A")["Total Shpenzimet"].sum()), ("Merdarë", ds.sheet("M")["Total Shpenzimet"].sum())]


def totals_subtitle(ds, data):
    return (f"n = {fmt(len(ds.sheet('A')))} (Aeroport) dhe {fmt(len(ds.sheet('M')))} (Merdarë) të anketuar   ·   "
            f"gjithsej {fmt(sum(v for _, v in data))} €")


def weights(ds):
    A, M = ds.sheet("A"), ds.sheet("M")
    return {"Pesha në mostër": [("Aeroport", len(A)), ("Merdarë", len(M))],
            "Pesha në shpenzime": [("Aeroport", A["Total Shpenzimet"].sum()), ("Merdarë", M["Total Shpenzimet"].sum())]}


CATEGORIES = [("para_arritjes", "Udhëtimi para mbërritjes"), ("Ushqim", "Ushqim"), ("blerjaProd", "Blerje e produkteve"),
              ("ParaFalur", "Para të falura familjarëve"), ("argetimi", "Argëtim"),
              ("shpenzimetTran", "Transporti në Kosovë"), ("vleraShpenzimit", "Shpenzime të tjera"),
              ("DhurataPer", "Dhurata për kosovarët"), ("akomodimi", "Akomodimi"),
              ("blerjePatun", "Blerje e patundshmërisë"), ("shpenzimeSh", "Shëndetësia")]


def structure_by_border(ds):
    """7.3: % of each category in each border point's total."""
    out = {}
    for name, sheet in [("Aeroport", "A"), ("Merdarë", "M")]:
        df = ds.sheet(sheet)
        s = {lab: pd.to_numeric(df[c], errors="coerce").sum() for c, lab in CATEGORIES}
        tot = sum(s.values())
        out[name] = {k: 100 * v / tot for k, v in s.items()}
    return out


def total_spending(ds):
    return sum(v for _, v in ds.table("Fig 3.1."))


# =============================================================================================  figures
FIGURAT = [
    # ---------------------------------------------------------------- 1. Struktura
    dict(file="fig_1_1_pika_kufitare.png", type="donut", table="Tabela 1.1.",
         title="Fig. 1.1. Të anketuarit sipas pikës kufitare", subtitle="n = {n} të anketuar"),
    dict(file="fig_1_2_grupi.png", type="donut", table="Tabela 1.2",
         title="Fig. 1.2. Ndarja e të anketuarve sipas grupit", subtitle="n = {n} të anketuar"),
    dict(file="fig_1_3_mosha.png", type="donut", table="Fig 1.2.", labels=["Mbi 18 vjeç", "Nën 18 vjeç"],
         title="Fig. 1.3. Personat mbi dhe nën moshën 18 vjeçare", subtitle="n = {n} persona"),
    dict(file="fig_1_4_shtetet_kryesore.png", type="donut", table="Fig 1.3.",
         title="Fig. 1.4. Të anketuarit sipas shteteve kryesore të vendbanimit", subtitle="n = {n} të anketuar"),
    dict(file="fig_1_5_te_gjitha_shtetet.png", type="hbar", table="Tabela 1.5.", value_format="pct_count",
         figsize=(10, 8.5), left=0.30,
         title="Fig. 1.5. Të anketuarit sipas të gjitha shteteve të vendbanimit", subtitle="n = {n} të anketuar"),
    dict(file="fig_1_6_pasaporta.png", type="donut", table="Tabela 1.6", colors=["accent", "navy"],
         labels=["Pasaporta të Kosovës", "Pasaporta të huaja"],
         title="Fig. 1.6. Personat sipas llojit të pasaportës", subtitle="n = {n} persona"),
    dict(file="fig_1_7_qellimi.png", type="donut", table="Tabela 1.7.",
         title="Fig. 1.7. Qëllimi i vizitës në Kosovë", subtitle="n = {n} të anketuar"),
    dict(file="fig_1_8_vizite_jashte.png", type="donut", table="Fig 1.5.",
         labels=["Vetëm në Kosovë", "Vizitë jashtë Kosovës"],
         title="Fig. 1.8. Qëndrimi vetëm në Kosovë kundrejt vizitave jashtë", subtitle="n = {n} të anketuar"),
    dict(file="fig_1_9_vendet_vizituara.png", type="hbar", table="Fig 1.6.", sort="desc", value_format="pct_count",
         figsize=(10, 4.2), left=0.30,
         title="Fig. 1.9. Vendet e vizituara jashtë Kosovës",
         subtitle="n = {n} të anketuar me vizitë jashtë Kosovës"),

    # ---------------------------------------------------------------- 2. Remitancat
    dict(file="fig_2_1_remitancat.png", type="donut", table="Fig 2.1.",
         title="Fig. 2.1. Dërgimi i remitancave në Kosovë", subtitle="n = {n} të anketuar"),
    dict(file="fig_2_2_format_remitancave.png", type="donut", table="Fig 2.2.", sort="desc",
         rename={"Agjenci për transfer parash": "Agjenci për\ntransfer parash"},
         colors=["navy", "accent", "blue", "gold", "teal"],
         title="Fig. 2.2. Format e dërgesave të remitancave",
         subtitle="n = {n} përmendje   ·   pyetje me zgjedhje të shumëfishtë"),
    dict(file="fig_2_3_formale_joformale.png", type="share_bar", table="Fig 2.3.",
         title="Fig. 2.3. Mënyrat formale dhe joformale të dërgimit",
         subtitle="n = {n} përmendje   ·   baza: zgjedhje, jo të anketuar"),

    # ---------------------------------------------------------------- 3. Shpenzimet
    dict(file="fig_3_1_shpenzimet_nen_10000.png", type="histogram", table="Tabela 3.2.", color="blue",
         title="Fig. 3.1. Shpenzimet e udhëtimit nën 10,000 €",
         subtitle=lambda ds, d: (f"n = {fmt(sum(v for _, v in d))} të anketuar   ·   mesatarja nën 10,000 €: "
                                 f"{fmt(ds.value('Mesatarja e shpenzimeve për intervistë, vetëm'))} €")),
    dict(file="fig_3_2_shpenzimet_mbi_10000.png", type="histogram", table="Tabela 3.3.", trim=True, color="blue",
         figsize=(8, 5.5),
         title="Fig. 3.2. Shpenzimet e udhëtimit mbi 10,000 €",
         subtitle="n = {n} të anketuar me shpenzim mbi 10,000 €"),
    dict(file="fig_3_3_veten_grupin.png", type="donut", table="Fig 3.1.", labels=["Individual", "Grupor"],
         title="Fig. 3.3. Përqindja e shpenzimeve për veten ose grupin",
         subtitle="njësia: euro të shpenzuara; gjithsej {n} €"),
    dict(file="fig_3_4_shpenzimet_shtetet.png", type="donut", table="Fig 3.2.",
         title="Fig. 3.4. Shpenzimet e ndara sipas shteteve kryesore",
         subtitle="njësia: euro të shpenzuara; gjithsej {n} €"),
    dict(file="fig_3_5_struktura_shpenzimeve.png", type="hbar", table="Aneks 2 / Tabela 1.", col=2, scale=100,
         sort="desc", value_format="pct", bold_values=True, figsize=(11, 7), left=0.40,
         title="Fig. 3.5. Struktura e shpenzimeve sipas kategorive",
         subtitle=lambda ds, d: (f"gjithsej {fmt(total_spending(ds))} €   ·   "
                                 "pjesa e secilës kategori në shpenzimet totale")),
    dict(file="fig_3_6_shendetesi.png", type="histogram", table="Tabela 3.4.", trim=True, color="teal",
         highlight_max=False, figsize=(10, 5.5),
         title="Fig. 3.6. Shpenzimet për shëndetësi",
         subtitle=lambda ds, d: (f"n = {fmt(sum(v for _, v in d))} të anketuar me shpenzim shëndetësor   ·   "
                                 f"gjithsej {fmt(ds.value('Shuma totale e shpenzimeve për shëndetësi'))} €")),
    dict(file="fig_3_7_nen_18.png", type="donut", table="Fig 3.3.", colors=["navy", "accent", "gold", "blue"],
         title="Fig. 3.7. Shpenzimet kryesore në grupet me persona nën 18 vjeç",
         subtitle=lambda ds, d: (f"n = {fmt(ds.value('Numri i intervistave me persona nën 18'))} intervista   ·   "
                                 "vetëm 4 kategori, jo shpenzimi total")),
    dict(file="fig_3_8_para_te_falura.png", type="histogram", table="Tabela 3.5.", trim=True, color="gold",
         title="Fig. 3.8. Para të falura familjarëve",
         subtitle="n = {n} të anketuar me para të falura"),

    # ---------------------------------------------------------------- 4–5
    dict(file="fig_4_1_pranimet.png", type="donut", calc=gifts, colors=["navy", "lightgrey"],
         title="Fig. 4.1. Pranimet nga Kosova",
         subtitle=lambda ds, d: (f"n = {fmt(sum(v for _, v in d))} të anketuar që iu përgjigjën pyetjes   ·   "
                                 f"vlera totale e dhuratave: {fmt(ds.value('Vlera totale e dhuratave të pranuara'))} €")),
    dict(file="fig_5_1_depozitat.png", type="donut", table="Fig 5.1.",
         title="Fig. 5.1. Depozitat e jo-rezidentëve në Kosovë", subtitle="n = {n} të anketuar"),
    dict(file="fig_5_2_format_depozitave.png", type="donut", table="Fig 5.2.",
         labels=["Para të gatshme", "Përmes bankës", "Në të dy format"],
         title="Fig. 5.2. Format e mbajtjes së depozitave", subtitle="n = {n} të anketuar me depozita"),

    # ---------------------------------------------------------------- 6. Pagesa dhe intensiteti (DB sheet)
    dict(file="fig_6_1_kartela.png", type="donut", calc=card, colors=["navy", "lightgrey"],
         title="Fig. 6.1. Përdorimi i kartelës për pagesa në Kosovë", subtitle=card_subtitle),
    dict(file="fig_6_2_ppd_gjithsej_pika.png", type="vbar", calc=ppd, colors=["teal", "navy", "accent"],
         ylabel="€ për person në ditë",
         title="Fig. 6.2. Shpenzimi për person në ditë",
         subtitle=lambda ds, d: f"n = {fmt(under_header(ds, 'Sa anketa perfshihen'))} intervista   ·   € për person në ditë"),
    dict(file="fig_6_3_ppd_shtetet.png", type="hbar", calc=ppd_by_country, color="teal", value_format="euro1",
         figsize=(10, 4.6), left=0.20, xmax_mult=1.2,
         title="Fig. 6.3. Shpenzimi për person në ditë, sipas shtetit",
         subtitle=f"shtete me së paku {MIN_INTERVIEWS} intervista   ·   € për person në ditë (në kllapa: numri i intervistave)"),

    # ---------------------------------------------------------------- 7. Pikat kufitare (A and M sheets)
    dict(file="fig_7_1_totalet_pika.png", type="vbar", calc=totals_by_border, colors=["navy", "accent"],
         value_format="euro", ylabel="Shpenzime totale (€)", width=0.45, left=0.14,
         title="Fig. 7.1. Shpenzimet totale sipas pikës kufitare", subtitle=totals_subtitle),
    dict(file="fig_7_2_pesha.png", type="share_rows", calc=weights, colors=["navy", "accent"],
         title="Fig. 7.2. Pesha në mostër kundrejt peshës në shpenzime",
         subtitle=lambda ds, d: (f"n = {fmt(sum(v for _, v in d['Pesha në mostër']))} të anketuar   ·   "
                                 f"shpenzimet: {fmt(sum(v for _, v in d['Pesha në shpenzime']))} €")),
    dict(file="fig_7_3_struktura_pika.png", type="grouped_hbar", calc=structure_by_border, colors=["navy", "accent"],
         figsize=(10, 8),
         title="Fig. 7.3. Struktura e shpenzimeve sipas pikës kufitare",
         subtitle="pjesa e secilës kategori në shpenzimet totale të pikës kufitare"),

    # ---------------------------------------------------------------- Aneks
    dict(file="fig_A1_remitancat_shtetet.png", type="hbar", table="Aneks 1 / Fig 1.", sort="desc", keep_last="Tjerat",
         value_format="pct_count", figsize=(10, 4.2), left=0.30,
         title="Fig. A1. Dërguesit e remitancave sipas shteteve", subtitle="n = {n} të anketuar që dërgojnë para"),
]


# =============================================================================================  run
def make_figures(dataset, dalja=None, viti=None, periudha="korrik–gusht", vetem=None):
    """Draw every report figure from the final workbook (Permbledhje + DB/A/M). Returns (folder, made, skipped)."""
    ds = bg.Dataset(dataset, relabel=RELABEL)
    year = viti or int(pd.to_datetime(ds.sheet("DB")["datapytesorit"]).dt.year.mode().iloc[0])
    out = dalja or f"figurat_{year}"
    source = f"Burimi: BQK — Anketa mbi shpenzimet e jo-rezidentëve në Kosovë, {periudha} {year}."
    print(f"Dataset: {dataset}   viti: {year}   dalja: {out}/")
    made, skipped = bg.run(FIGURAT, ds, out, source, only=vetem)
    return out, made, skipped


def main():
    ap = argparse.ArgumentParser(description="Make all report figures from the survey dataset.")
    ap.add_argument("dataset", help="survey workbook (.xlsx) with the Permbledhje sheet")
    ap.add_argument("--dalja", help="output folder (default: figurat_<year>)")
    ap.add_argument("--viti", type=int, help="survey year (default: from interview dates in DB)")
    ap.add_argument("--periudha", default="korrik–gusht", help="survey period shown in the source line")
    ap.add_argument("--vetem", nargs="*", help="only figures whose file name or title starts with these")
    a = ap.parse_args()
    out, made, skipped = make_figures(**vars(a))
    print(f"\n{len(made)} figura të krijuara" + (f", {len(skipped)} të anashkaluara:" if skipped else "."))
    for f, why in skipped:
        print("   -", f, "→", why)


if __name__ == "__main__":
    main()
