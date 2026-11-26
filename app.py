"""
app.py — upload the raw survey export, get all report charts.

    python -m streamlit run app.py

Runs on your own computer: the files are processed in a temporary folder and deleted afterwards; nothing is sent
anywhere or kept. The data is cleaned on the way (same rules as the report workbook); only the charts come out.
"""
import io
import os
import tempfile

import streamlit as st

import bqk

st.set_page_config(page_title="BQK – Grafikët e jo-rezidentëve", layout="wide")
st.title("Grafikët – shpenzimet e jo-rezidentëve në Kosovë")
st.caption("Ngarkoni datasetin bruto → merrni të gjithë grafikët e raportit.")

with st.form("hyrja"):
    raw = st.file_uploader("Dataseti bruto (eksporti i anketës, .xlsx ose .csv)", type=["xlsx", "xls", "csv"])
    c1, c2 = st.columns(2)
    codebook = c1.file_uploader("Codebook (opsional, rekomandohet për vitet e vjetra)", type=["xlsx"])
    korr = c2.file_uploader("Korrigjime manuale (opsional, CSV)", type=["csv"])
    c3, c4, c5, c6 = st.columns(4)
    viti = c3.number_input("Viti (0 = nga datat e intervistave)", min_value=0, max_value=2100, value=0, step=1)
    outliers = c4.text_input("Outlier-at (referues, të ndarë me presje)", "")
    pragu = c5.number_input("Pragu i shpenzimeve (€)", min_value=0, value=10000, step=500)
    periudha = c6.text_input("Periudha e anketës", "korrik–gusht")
    pika = st.selectbox("Pika kufitare – vetëm nëse eksporti NUK ka kolonën e pikës (skedar vetëm për një pikë)",
                        ["Eksporti e ka kolonën", "Aeroport", "Merdarë"])
    go = st.form_submit_button("Krijo grafikët", type="primary")


def save(upload, folder):
    path = os.path.join(folder, upload.name)
    with open(path, "wb") as f:
        f.write(upload.getbuffer())
    return path


if go:
    if raw is None:
        st.error("Ngarkoni datasetin bruto.")
        st.stop()
    try:
        ol = [int(x) for x in outliers.replace(";", ",").replace(" ", ",").split(",") if x.strip()]
    except ValueError:
        st.error("Outlier-at duhet të jenë numra (referues), p.sh. 52, 144, 147")
        st.stop()
    with tempfile.TemporaryDirectory() as tmp, st.spinner("Duke përgatitur të dhënat dhe duke vizatuar grafikët…"):
        try:
            r, out, made, skipped, zpath = bqk.run(
                save(raw, tmp), dalja=os.path.join(tmp, "grafiket"),
                codebook=save(codebook, tmp) if codebook else None,
                korrigjime=save(korr, tmp) if korr else None,
                viti=int(viti) or None, outliers=ol, pragu=float(pragu), periudha=periudha,
                pika={"Aeroport": "A", "Merdarë": "M"}.get(pika))
        except SystemExit as e:                     # e.g. the file is not a survey export
            st.error(str(e))
            st.stop()
        except Exception as e:
            st.exception(e)
            st.stop()
        # keep the charts in memory so the download buttons work after the temp folder is gone
        st.session_state["rez"] = {
            "r": r, "skipped": skipped,
            "zip": open(zpath, "rb").read(), "zipname": os.path.basename(zpath),
            "figs": [(f, open(os.path.join(out, f), "rb").read()) for f in made],
        }

rez = st.session_state.get("rez")
if rez:
    r = rez["r"]
    st.success(f"Viti {r['viti']}: {len(rez['figs'])} grafikë nga {r['n_db']} intervista të vlefshme "
               f"({r['n_A']} Aeroport, {r['n_M']} Merdarë"
               + "".join(f", {v} {k}" for k, v in r["pikat_tjera"].items()) + ")")
    if any(w[1] == "pytesorivendi" and "nuk ka kolonën" in w[4] for w in r["paralajmerime"]):
        st.warning("Ky skedar nuk ka kolonën e pikës kufitare (pytesorivendi), ndaj grafikët sipas Aeroport/Merdarë "
                   "dalin bosh. Kontrolloni a është eksporti i plotë; nëse skedari është vetëm për një pikë, "
                   "zgjidhni pikën më lart dhe provoni përsëri.")
    st.download_button("⬇ Shkarko të gjithë grafikët (.zip)", rez["zip"], rez["zipname"], "application/zip",
                       type="primary")
    if r["paralajmerime"]:
        with st.expander(f"⚠ {len(r['paralajmerime'])} paralajmërime për kontroll"):
            st.table([{"kolona": w[1], "paralajmërimi": w[4]} for w in r["paralajmerime"]])
    if rez["skipped"]:
        with st.expander(f"Grafikë të anashkaluar ({len(rez['skipped'])})"):
            for f, why in rez["skipped"]:
                st.write(f"- **{f}**: {why}")

    cols = st.columns(2)
    for i, (name, png) in enumerate(rez["figs"]):
        with cols[i % 2]:
            st.image(io.BytesIO(png), caption=name, width="stretch")
            st.download_button("⬇ " + name, png, name, "image/png", key="dl_" + name)
