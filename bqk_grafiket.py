"""
bqk_grafiket.py — chart template for the BQK non-resident spending survey report.

Holds everything that should look the same every year: the palette, fonts, title/subtitle/source layout,
and the chart types used in the report. It knows nothing about specific figures; those are listed in
figurat.py, which reads the data and calls `run()`.

Chart types (the "type" field of a figure):
    donut        ring chart with total in the centre and labelled slices        data: [(label, value), ...]
    hbar         horizontal bars, one colour                                    data: [(label, value), ...]
    histogram    vertical bars over intervals, tallest bar highlighted          data: [(interval, count), ...]
    vbar         vertical bars, one colour per bar (e.g. Total / Airport / ...) data: [(label, value), ...]
    share_bar    one 100 % stacked bar (e.g. formal vs informal)               data: [(label, value), ...]
    share_rows   several 100 % stacked bars, one per row                        data: {row: [(series, value), ...]}
    grouped_hbar horizontal bars, several series side by side                   data: {series: {category: value}}
"""
import os
import warnings
import numpy as np
import pandas as pd
import openpyxl
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter

PALETTE = {
    "navy": "#1F3864",
    "blue": "#2E75B6",
    "lightblue": "#9DC3E6",
    "paleblue": "#D6E4F0",
    "accent": "#C55A11",
    "accent_light": "#F4B183",
    "gold": "#BF9000",
    "teal": "#2F7A78",
    "grey": "#808080",
    "lightgrey": "#D9D9D9",
    "text": "#262626",
}
P = PALETTE
DEFAULT_COLORS = ["navy", "accent", "blue", "gold", "teal", "lightblue", "accent_light"]

plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "text.color": P["text"],
    "axes.labelcolor": P["text"],
    "xtick.color": P["text"],
    "ytick.color": P["text"],
    "axes.edgecolor": P["grey"],
    "savefig.dpi": 200,
})


# =============================================================================================  formatting
def fmt(x, dec=0):
    """1234567.8 -> '1,234,568' (dec=0) or '1,234,567.80' (dec=2)."""
    return f"{x:,.{dec}f}"


def color(c):
    """Palette name ('navy') or hex ('#1F3864') -> hex."""
    return P.get(c, c)


VALUE_FORMATS = {
    "count": lambda v, tot: fmt(v),
    "pct_count": lambda v, tot: f"{100 * v / tot:.1f}%  ({fmt(v)})",
    "pct": lambda v, tot: f"{v:.1f}%",                 # value already in percent
    "share": lambda v, tot: f"{100 * v:.1f}%",         # value is a share 0–1
    "euro": lambda v, tot: f"{fmt(v)} €",
    "euro1": lambda v, tot: f"{v:,.1f} €",
}


# =============================================================================================  data access
class Dataset:
    """
    Reads the survey workbook.
      * the summary sheet ("Permbledhje") is searched by table TITLE, not by cell address,
        so tables can move between years without breaking anything;
      * other sheets (DB, A, M, …) are available as pandas DataFrames via .sheet(name).
    `relabel` renames labels coming from the sheet (e.g. codebook names for unlabelled codes).
    """

    def __init__(self, xlsx, summary_sheet="Permbledhje", relabel=None):
        self.xlsx = xlsx
        self.relabel = relabel or {}
        ws = openpyxl.load_workbook(xlsx, data_only=True)[summary_sheet]
        self.rows = [list(r) for r in ws.iter_rows(values_only=True)]
        self._sheets = {}

    def find(self, title):
        """Row index of the first row whose first cell starts with `title`."""
        for i, r in enumerate(self.rows):
            if isinstance(r[0], str) and r[0].strip().startswith(title):
                return i
        raise KeyError(f"No table titled '{title}' in the summary sheet")

    def table(self, title, col=1, total=False):
        """
        [(label, value), ...] for the table under `title`: skips note rows, takes the header row,
        then reads until an empty row or the 'Total' row. col=1 → counts/values, col=2 → shares.
        """
        i = self.find(title) + 1
        while not (self.rows[i][1] is not None and isinstance(self.rows[i][1], str)):   # header row
            i += 1
        out = []
        for r in self.rows[i + 1:]:
            if r[0] is None:
                break
            is_total = str(r[0]).startswith("Total")
            if is_total and not total:
                break
            out.append((self.relabel.get(str(r[0]), str(r[0])), r[col]))
            if is_total:
                break
        return out

    def value(self, title, col=1):
        """Single value on the row whose label starts with `title` (e.g. 'Shuma totale e …')."""
        return self.rows[self.find(title)][col]

    def row(self, title):
        """All non-empty cells of the row starting with `title` (label included)."""
        return [v for v in self.rows[self.find(title)] if v is not None]

    def sheet(self, name):
        if name not in self._sheets:
            self._sheets[name] = pd.read_excel(self.xlsx, sheet_name=name)
        return self._sheets[name]


# =============================================================================================  layout
def _header(fig, title, subtitle, size):
    H = fig.get_figheight()
    fig.text(0.02, 1 - 0.22 / H, title, fontsize=size, fontweight="bold", color=P["navy"], va="top")
    if subtitle:
        fig.text(0.02, 1 - (0.30 + size / 72 * 1.25) / H, subtitle, fontsize=size * 0.68, color=P["grey"], va="top")


def _top(fig, inches=1.05):
    return 1 - inches / fig.get_figheight()


def _footer(fig, text):
    fig.text(0.02, 0.025, text, fontsize=8.5, color=P["grey"], va="bottom")


def _clean(ax, keep=("left",)):
    for s in ["top", "right", "bottom", "left"]:
        if s not in keep:
            ax.spines[s].set_visible(False)


def _colors(spec, n):
    cols = spec.get("colors") or DEFAULT_COLORS
    return [color(cols[i % len(cols)]) for i in range(n)]


# =============================================================================================  chart types
def donut(spec, data, source):
    labels, values = zip(*data)
    fig = plt.figure(figsize=spec.get("figsize", (10, 7.5)))
    _header(fig, spec["title"], spec.get("subtitle"), 15)
    ax = fig.add_axes([0.22, 0.12, 0.56, 0.72])
    total = sum(values)
    wedges, _ = ax.pie(values, colors=_colors(spec, len(values)), startangle=90, counterclock=False,
                       wedgeprops=dict(width=0.38, edgecolor="white", linewidth=2.5))
    ax.text(0, 0.08, spec.get("center", fmt(total)), ha="center", va="center", fontsize=24, fontweight="bold",
            color=P["navy"])
    ax.text(0, -0.17, spec.get("center_label", "gjithsej"), ha="center", va="center", fontsize=12, color=P["grey"])
    for w, lab, v in zip(wedges, labels, values):
        ang = np.deg2rad((w.theta2 + w.theta1) / 2)
        x, y = np.cos(ang), np.sin(ang)
        ax.annotate(f"{lab}\n{100 * v / total:.1f}%  ({fmt(v)})",
                    xy=(0.81 * x, 0.81 * y), xytext=(1.25 * x + (0.05 if x >= 0 else -0.05), 1.22 * y),
                    ha="left" if x >= 0 else "right", va="center", fontsize=12.5,
                    arrowprops=dict(arrowstyle="-", color=P["grey"], lw=0.8))
    ax.set_aspect("equal")
    _footer(fig, source)
    return fig


def hbar(spec, data, source):
    labels, values = [d[0] for d in data], [d[1] for d in data]
    notes = [d[2] if len(d) > 2 else None for d in data]      # optional 3rd item, shown as "(note)"
    n = len(values)
    figsize = spec.get("figsize", (10, max(4.2, 2.4 + 0.3 * n)))
    left = spec.get("left", min(0.45, 0.10 + 0.0105 * max(len(str(l)) for l in labels)))
    vf = spec.get("value_format", "count")
    vf = VALUE_FORMATS[vf] if isinstance(vf, str) else vf
    total = spec.get("base", sum(values))
    bold = spec.get("bold_values", False)
    fig = plt.figure(figsize=figsize)
    _header(fig, spec["title"], spec.get("subtitle"), 13)
    b0 = 0.55 / figsize[1]
    ax = fig.add_axes([left, b0, 0.92 - left, _top(fig) - b0])
    y = np.arange(n)
    ax.barh(y, values, color=color(spec.get("color", "navy")), height=0.6)
    ax.set_yticks(y, labels, fontsize=10.5)
    ax.invert_yaxis()
    vmax = max(values)
    for yi, v, note in zip(y, values, notes):
        ax.text(v + vmax * 0.012, yi, vf(v, total) + (f"  ({note})" if note is not None else ""), va="center", fontsize=10,
                color=P["navy"] if bold else P["text"], fontweight="bold" if bold else "normal")
    ax.set_xlim(0, vmax * spec.get("xmax_mult", 1.25))
    ax.xaxis.set_visible(False)
    _clean(ax)
    ax.spines["left"].set_color(P["lightgrey"])
    ax.tick_params(axis="y", length=0)
    _footer(fig, source)
    return fig


def histogram(spec, data, source):
    bins, counts = zip(*data)
    fig = plt.figure(figsize=spec.get("figsize", (11, 6)))
    _header(fig, spec["title"], spec.get("subtitle"), 13)
    ax = fig.add_axes([0.08, 0.24, 0.90, 0.60])
    x = np.arange(len(bins))
    cols = [color(spec.get("color", "blue"))] * len(bins)
    if spec.get("highlight_max", True):
        cols[int(np.argmax(counts))] = P["navy"]
    ax.bar(x, counts, color=cols, width=0.72)
    for xi, c in zip(x, counts):
        if c > 0:
            ax.text(xi, c + max(counts) * 0.015, fmt(c), ha="center", va="bottom", fontsize=9)
    ax.set_xticks(x, bins, rotation=45, ha="right", fontsize=9)
    ax.set_ylabel(spec.get("ylabel", "Numri i të anketuarve"), fontsize=10)
    ax.set_ylim(0, max(counts) * 1.12)
    ax.grid(axis="y", color=P["lightgrey"], lw=0.7)
    ax.set_axisbelow(True)
    _clean(ax, keep=("left", "bottom"))
    _footer(fig, source)
    return fig


def vbar(spec, data, source):
    labels, values = zip(*data)
    vf = spec.get("value_format", "euro1")
    vf = VALUE_FORMATS[vf] if isinstance(vf, str) else vf
    fig = plt.figure(figsize=spec.get("figsize", (9, 6)))
    _header(fig, spec["title"], spec.get("subtitle"), 14)
    ax = fig.add_axes([spec.get("left", 0.10), 0.12, 0.96 - spec.get("left", 0.10), _top(fig, 1.15) - 0.12])
    x = np.arange(len(values))
    ax.bar(x, values, color=_colors(spec, len(values)), width=spec.get("width", 0.55))
    for xi, v in zip(x, values):
        ax.text(xi, v + max(values) * 0.012, vf(v, sum(values)), ha="center", va="bottom", fontsize=12,
                fontweight="bold")
    ax.set_xticks(x, labels, fontsize=11.5)
    ax.set_ylim(0, max(values) * 1.15)
    ax.set_ylabel(spec.get("ylabel", ""), fontsize=10)
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: fmt(v)))
    ax.grid(axis="y", color=P["lightgrey"], lw=0.7)
    ax.set_axisbelow(True)
    _clean(ax, keep=("left", "bottom"))
    _footer(fig, source)
    return fig


def share_bar(spec, data, source):
    labels, values = zip(*data)
    tot = sum(values)
    fig = plt.figure(figsize=spec.get("figsize", (10, 3.8)))
    _header(fig, spec["title"], spec.get("subtitle"), 13)
    ax = fig.add_axes([0.03, 0.18, 0.94, _top(fig, 1.2) - 0.18])
    left = 0
    for lab, v, c in zip(labels, values, _colors(spec, len(values))):
        ax.barh(0, v / tot, left=left, color=c, height=0.8)
        ax.text(left + v / tot / 2, 0, f"{lab}\n{100 * v / tot:.1f}%", ha="center", va="center", color="white",
                fontweight="bold", fontsize=11)
        left += v / tot
    ax.set_xlim(0, 1)
    ax.set_yticks([])
    ax.set_xticks(np.linspace(0, 1, 6), [f"{int(t * 100)}%" for t in np.linspace(0, 1, 6)])
    _clean(ax, keep=("bottom",))
    _footer(fig, source)
    return fig


def share_rows(spec, data, source):
    rows = list(data)
    series = [s for s, _ in data[rows[0]]]
    cols = _colors(spec, len(series))
    fig = plt.figure(figsize=spec.get("figsize", (10, 4.6)))
    _header(fig, spec["title"], spec.get("subtitle"), 13)
    ax = fig.add_axes([0.20, 0.15, 0.75, _top(fig, 1.4) - 0.15])
    for yi, r in enumerate(rows):
        vals = [v for _, v in data[r]]
        tot, left = sum(vals), 0
        for si, v in enumerate(vals):
            ax.barh(yi, v / tot, left=left, color=cols[si], height=0.6, label=series[si] if yi == 0 else None)
            ax.text(left + v / tot / 2, yi, f"{100 * v / tot:.1f}%", ha="center", va="center", color="white",
                    fontweight="bold", fontsize=11)
            left += v / tot
    ax.set_yticks(range(len(rows)), rows, fontsize=10.5)
    ax.invert_yaxis()
    ax.set_xlim(0, 1)
    ax.set_xticks(np.linspace(0, 1, 6), [f"{int(t * 100)}%" for t in np.linspace(0, 1, 6)])
    _clean(ax, keep=("bottom",))
    ax.tick_params(axis="y", length=0)
    ax.legend(loc="lower left", bbox_to_anchor=(0, 1.0), ncol=len(series), frameon=False, fontsize=10)
    _footer(fig, source)
    return fig


def grouped_hbar(spec, data, source):
    series = list(data)
    cats = list(data[series[0]])
    if spec.get("sort", True):
        cats = sorted(cats, key=lambda k: -data[series[0]][k])
    cols = _colors(spec, len(series))
    fig = plt.figure(figsize=spec.get("figsize", (10, max(5, 1.8 + 0.55 * len(cats)))))
    _header(fig, spec["title"], spec.get("subtitle"), 13)
    ax = fig.add_axes([spec.get("left", 0.28), 0.10, 0.94 - spec.get("left", 0.28), _top(fig, 1.2) - 0.10])
    y = np.arange(len(cats))
    h = 0.76 / len(series)
    vmax = max(max(d.values()) for d in data.values())
    for si, s in enumerate(series):
        off = (si - (len(series) - 1) / 2) * h
        vals = [data[s][k] for k in cats]
        ax.barh(y + off, vals, height=h, color=cols[si], label=s)
        for yi, v in zip(y, vals):
            ax.text(v + vmax * 0.01, yi + off, spec.get("value_fmt", "{:.1f}%").format(v), va="center",
                    fontsize=9.5, color=cols[si], fontweight="bold")
    ax.set_yticks(y, cats, fontsize=10.5)
    ax.invert_yaxis()
    ax.set_xlim(0, vmax * 1.15)
    ax.xaxis.set_visible(False)
    _clean(ax)
    ax.tick_params(axis="y", length=0)
    ax.legend(loc="lower right", frameon=False, fontsize=10)
    _footer(fig, source)
    return fig


CHART_TYPES = {"donut": donut, "hbar": hbar, "histogram": histogram, "vbar": vbar,
               "share_bar": share_bar, "share_rows": share_rows, "grouped_hbar": grouped_hbar}


# =============================================================================================  data shaping
def prepare(spec, ds):
    """Turn a figure entry into chart data: read the table or call `calc`, then apply options."""
    if "calc" in spec:
        data = spec["calc"](ds)
    else:
        data = ds.table(spec["table"], col=spec.get("col", 1))
    if isinstance(data, dict):            # multi-series charts: returned as-is
        return data
    data = [tuple(d) for d in data]
    if "labels" in spec:                  # replace labels in order
        data = [(lab,) + d[1:] for lab, d in zip(spec["labels"], data)]
    if "rename" in spec:
        data = [(spec["rename"].get(d[0], d[0]),) + d[1:] for d in data]
    if spec.get("scale"):
        data = [(d[0], d[1] * spec["scale"]) + d[2:] for d in data]
    if spec.get("sort") == "desc":
        last = spec.get("keep_last")      # e.g. 'Tjerat' stays at the bottom
        data = sorted(data, key=lambda kv: (kv[0] == last, -kv[1]))
    if spec.get("trim"):                  # drop empty intervals at the end
        nz = [i for i, d in enumerate(data) if d[1]]
        data = data[:max(nz) + 1]
    return data


def subtitle(spec, data, ds):
    sub = spec.get("subtitle")
    if callable(sub):
        return sub(ds, data)
    if isinstance(sub, str) and isinstance(data, list):
        return sub.format(n=fmt(sum(d[1] for d in data)))
    return sub


def run(figures, ds, outdir, source, only=None):
    """
    Draw every figure in `figures` and save it to `outdir`.
    A figure whose table is missing or whose data fails is skipped with a warning; the rest still run.
    Returns (made, skipped).
    """
    os.makedirs(outdir, exist_ok=True)
    made, skipped = [], []
    for spec in figures:
        if only and not any(spec["file"].startswith(o) or spec["title"].startswith(o) for o in only):
            continue
        try:
            data = prepare(spec, ds)
            spec = dict(spec, subtitle=subtitle(spec, data, ds))
            fig = CHART_TYPES[spec["type"]](spec, data, spec.get("source", source))
            fig.savefig(os.path.join(outdir, spec["file"]), facecolor="white")
            plt.close(fig)
            made.append(spec["file"])
            print("  ✓", spec["file"])
        except Exception as e:                 # keep going; report at the end
            plt.close("all")
            skipped.append((spec["file"], f"{type(e).__name__}: {e}"))
            print("  ✗", spec["file"], "—", e)
    return made, skipped
