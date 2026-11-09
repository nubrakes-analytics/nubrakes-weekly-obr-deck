"""
NuBrakes Weekly Operations Business Review — render layer.

Reads metrics.json, draws the charts, writes the PowerPoint. The deck structure
is fixed week to week: four slides, same layout, same tables, same charts. Only
the numbers and the auto-drafted callouts change.

Usage:
    python render.py [--metrics out/metrics.json] [--out out]
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt, Emu

# --- Brand ------------------------------------------------------------------
NAVY = "#00316D"        # VIOC Dark Blue  (Pantone 280)
BLUE = "#078BC7"        # VIOC Light Blue (Pantone Process Blue)
RED = "#E51C24"         # VIOC Red        (Pantone 485)
GRAY = "#999999"        # VIOC Gray       (Pantone 877)
BLACK = "#000000"
GREEN = "#1B7F4B"       # positive-variance ink only; not a chart series colour
INK = "#1A1A1A"
MUTED = "#6B7280"
RULE = "#D8DDE3"
BAND = "#F4F6F8"
SURFACE = "#FFFFFF"

FONT = "Calibri"
CHART_FONT = "Carlito"  # metric-compatible with Calibri

plt.rcParams.update({
    "font.family": CHART_FONT, "font.size": 9,
    "axes.edgecolor": RULE, "axes.linewidth": 0.8, "axes.labelcolor": MUTED,
    "xtick.color": MUTED, "ytick.color": MUTED,
    "xtick.labelsize": 8, "ytick.labelsize": 8,
    "figure.facecolor": SURFACE, "axes.facecolor": SURFACE,
    "savefig.facecolor": SURFACE,
})

SLIDE_W, SLIDE_H = 13.333, 7.5
MARGIN = 0.42
CONTENT_W = SLIDE_W - 2 * MARGIN

# Profitability & unit economics is out of the weekly deck for now. The compute layer
# still loads the EGP model and the labor file, so this is the only switch to flip.
INCLUDE_PROFITABILITY = False

# Heat-map tints for the technician leaderboard. Rank-based within a market, so a
# strong technician in a weak market still reads as strong for that market.
HEAT = ["#FBDDDF", "#FDEEEF", None, "#EAF6EF", "#D7EDDF"]

# Leaderboard columns. `shade` marks the rank-scored ones; volume-driven columns
# (days, revenue, jobs, hours) stay unshaded — ranking them would just re-rank headcount.
LEADERBOARD = [
    ("days", "Days", "days", False, True),
    ("revenue", "Revenue", "money0", False, True),
    ("jobs", "Rev jobs", "int", False, True),
    ("jobs_per_day", "Rev jobs/day", "num1", True, True),
    # Booked ticket is what the technician was handed, not what they did with it, so it
    # is left unshaded alongside the volume columns. Ranking a technician on the value
    # of the work dispatched to them would score the schedule, not the person.
    ("booked_ticket", "Booked ticket", "money0", False, True),
    # Booked ticket, average ticket and ticket completion sit together: handed, closed,
    # and the spread between them, read left to right.
    ("avg_ticket", "Avg ticket", "money0", True, True),
    ("ticket_completion", "Ticket compl.", "pct1", True, True),
    ("service_call_rate", "Svc call", "pct1", True, False),
    ("rotor_attach_rate", "Rotor attach", "pct1", True, True),
    ("invoiced_hours", "Inv. hours", "hours", False, True),
    ("efficiency", "Efficiency", "pct0", True, True),
]


def rgb(hex_str):
    return RGBColor.from_string(hex_str.lstrip("#").upper())


# --- Formatters -------------------------------------------------------------
def money(v, dp=0):
    if v is None:
        return "—"
    return f"${v:,.{dp}f}" if v >= 0 else f"-${abs(v):,.{dp}f}"


def money_k(v):
    if v is None:
        return "—"
    sign = "-" if v < 0 else ""
    a = abs(v)
    return f"{sign}${a/1000:,.0f}k" if a >= 1000 else f"{sign}${a:,.0f}"


def pct(v, dp=1):
    return "—" if v is None else f"{v*100:.{dp}f}%"


def signed_pct(v, dp=1):
    return "—" if v is None else f"{v*100:+.{dp}f}%"


def num(v, dp=2):
    return "—" if v is None else f"{v:,.{dp}f}"


def pct_delta(cur, prior):
    if cur is None or prior in (None, 0):
        return None
    return (cur - prior) / abs(prior)


def num1(v):
    return num(v, 1)


def num2(v):
    return num(v, 2)


def pct1(v):
    return pct(v, 1)


def money0(v):
    return money(v, 0)


def as_int(v):
    return "—" if v is None else f"{v:,.0f}"


def hours(v):
    return "—" if v is None else f"{v:,.0f}h"


def pct0(v):
    return pct(v, 0)


FORMATTERS = {"money0": money0, "num1": num1, "num2": num2, "pct1": pct1,
              "pct0": pct0, "int": as_int, "days": as_int, "hours": hours}


def fmt_date(iso):
    d = dt.date.fromisoformat(iso)
    return d.strftime("%b %-d")


def month_label(iso_month):
    return dt.date.fromisoformat(iso_month + "-01").strftime("%b")


# --- Charts -----------------------------------------------------------------
def chart_sparkline(values, path, good_up=True, width=2.70, height=0.44):
    """The strip inside a KPI tile. No axes — it carries shape, not magnitude; the
    tile's own value and delta carry magnitude."""
    vals = [v for v in values if v is not None]
    if len(vals) < 3:
        vals = vals or [0]
    fig, ax = plt.subplots(figsize=(width, height), dpi=200)
    x = range(len(vals))
    ax.fill_between(x, vals, min(vals) - (max(vals) - min(vals) or 1) * 0.25,
                    color=BLUE, alpha=0.14, zorder=2)
    ax.plot(x, vals, color=BLUE, lw=1.3, zorder=3, solid_capstyle="round")
    last_up = len(vals) > 1 and vals[-1] >= vals[-2]
    end = GREEN if (last_up == good_up) else RED
    ax.plot([len(vals) - 1], [vals[-1]], "o", ms=3.4, color=end, mec=SURFACE, mew=0.9,
            zorder=4)
    ax.set_ylim(min(vals) - (max(vals) - min(vals) or 1) * 0.25,
                max(vals) + (max(vals) - min(vals) or 1) * 0.18)
    ax.axis("off")
    fig.subplots_adjust(left=0, right=1, top=1, bottom=0)
    fig.savefig(path, transparent=True)
    plt.close(fig)


def chart_month_pace(m, path, width=6.05, height=2.85):
    """Closed months as solid bars, the month in flight as its month-to-date actual
    with the straight-line pace stacked on in outline. The comparison Walker reads is
    the pace bar against the previous month's bar."""
    s = m["monthly"]["series"]
    labels = [x["label"] for x in s]
    fig, ax = plt.subplots(figsize=(width, height), dpi=200)
    x = list(range(len(s)))

    closed = [x_["total"] if not x_["partial"] else 0 for x_ in s]
    ax.bar(x, closed, color=BLUE, width=0.62, zorder=3)

    prior_closed = None
    for i, e in enumerate(s):
        if e["partial"]:
            ax.bar([i], [e["total"]], color=NAVY, width=0.62, zorder=4)
            ax.bar([i], [e["pace"] - e["total"]], bottom=[e["total"]], width=0.62,
                   facecolor="none", edgecolor=NAVY, lw=1.1, ls=(0, (3, 2)), zorder=4)
            ax.annotate(f"{e['pace']/1000:,.0f}\npace", (i, e["pace"]),
                        textcoords="offset points", xytext=(0, 4), ha="center",
                        fontsize=7, color=NAVY, fontweight="bold", linespacing=1.15)
            ax.annotate(f"MTD\n{e['total']/1000:,.0f}", (i, e["total"] / 2), ha="center",
                        va="center", fontsize=7.4, color="#FFFFFF", fontweight="bold",
                        linespacing=1.1)
            if prior_closed:
                d = (e["pace"] - prior_closed) / prior_closed
                ax.annotate(f"{d*100:+.1f}% vs {s[i-1]['label']}",
                            (i, e["pace"]), textcoords="offset points",
                            xytext=(0, 26), ha="center", fontsize=7.4,
                            color=GREEN if d >= 0 else RED, fontweight="bold")
        else:
            ax.annotate(f"{e['total']/1000:,.0f}", (i, e["total"]),
                        textcoords="offset points", xytext=(0, 3), ha="center",
                        fontsize=7, color=MUTED)
            prior_closed = e["total"]

    top = max(max(e["total"] for e in s), max(e["pace"] or 0 for e in s))
    ax.set_ylim(0, top * 1.30)
    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontsize=8)
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"${v/1000:,.0f}k"))
    ax.grid(axis="y", color=RULE, lw=0.6, zorder=0)
    ax.set_axisbelow(True)
    for sp in ("top", "right", "left"):
        ax.spines[sp].set_visible(False)
    cur_name = dt.date.fromisoformat(m["monthly"]["current_month"] + "-01").strftime("%B")
    ax.set_title(f"Net revenue by month — {cur_name} paced from day "
                 f"{m['monthly']['days_elapsed']} of {m['monthly']['days_in_month']}",
                 fontsize=9.5, color=NAVY, fontweight="bold", loc="left", pad=8)
    fig.tight_layout(pad=0.6)
    fig.savefig(path)
    plt.close(fig)


def chart_region_trend(m, region, metric, title, path, width=4.07, height=2.30,
                       fmt=None, good_up=True):
    """Region aggregate in navy over its component markets in thin blue. Two series
    identities only — the market lines show spread, not identity, so no 6-hue palette."""
    weeks = [dt.date.fromisoformat(w) for w in m["trend_weeks"]]
    markets = m["regions"][region]["markets"]
    series = m["trends"][metric]
    fig, ax = plt.subplots(figsize=(width, height), dpi=200)
    x = range(len(weeks))

    reg = series[f"region:{region}"]
    clean = [v for v in reg if v is not None]
    ax.fill_between(x, reg, min(clean) * 0.90 if clean else 0, color=BLUE, alpha=0.12,
                    zorder=2)
    ax.plot(x, reg, color=NAVY, lw=1.9, zorder=4, solid_capstyle="round",
            marker="o", ms=2.8, mfc=NAVY, mec=SURFACE, mew=0.8)

    # Every point labelled — the region line is the only series, so there is room.
    # Each label sits on the outside of the curve: above a peak, below a trough. That
    # keeps it clear of the line itself, which blind alternation does not on a noisy
    # series.
    for i, v in enumerate(reg):
        if v is None:
            continue
        last = i == len(reg) - 1
        nbrs = [reg[j] for j in (i - 1, i + 1) if 0 <= j < len(reg) and reg[j] is not None]
        above = True if not nbrs else v >= sum(nbrs) / len(nbrs)
        ax.annotate(fmt(v) if fmt else num(v), (i, v), textcoords="offset points",
                    xytext=(0, 8 if above else -13), ha="center",
                    fontsize=7.6 if last else 6.5,
                    color=NAVY if last else MUTED,
                    fontweight="bold" if last else "normal", zorder=6)

    ax.set_xticks(list(x))
    ax.set_xticklabels([w.strftime("%-m/%-d") for w in weeks], fontsize=6.2)
    ax.set_xlim(-0.7, len(reg) - 0.3)   # room for the end labels
    ax.tick_params(axis="y", labelsize=7)
    if clean:
        span = (max(clean) - min(clean)) or (abs(max(clean)) * 0.1 or 1)
        ax.set_ylim(min(clean) - span * 0.75, max(clean) + span * 0.55)
    if fmt is pct1:
        ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v*100:.0f}%"))
    elif fmt is money0:
        ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"${v:,.0f}"))
    ax.grid(axis="y", color=RULE, lw=0.6, zorder=0)
    ax.set_axisbelow(True)
    for sp in ("top", "right", "left"):
        ax.spines[sp].set_visible(False)
    ax.set_title(f"{title} — {region} overall", fontsize=8.6, color=NAVY,
                 fontweight="bold", loc="left", pad=6)
    fig.tight_layout(pad=0.5)
    fig.savefig(path)
    plt.close(fig)


def chart_market_panel(m, market, path, width=5.30, height=2.66):
    """Four mini trends for one market — the same metric set as the exec tiles, so a
    manager reads the market slide in the language of the company slide."""
    weeks = [dt.date.fromisoformat(w) for w in m["trend_weeks"]]
    # Compact label forms: the panel titles and axes carry the units, so the labels
    # only need the figure. Thirteen of them per panel otherwise will not fit.
    def k_short(v):
        return "—" if v is None else f"{v/1000:.0f}"

    def d_short(v):
        return "—" if v is None else f"{v:,.0f}"

    def n_short(v):
        return "—" if v is None else f"{v:.1f}"

    def p_short(v):
        return "—" if v is None else f"{v*100:.0f}"

    specs = [("revenue", "Net revenue ($000)", money_k, k_short, True),
             ("avg_ticket", "Average ticket ($)", money0, d_short, True),
             ("jobs_per_tech_day", "Rev-producing jobs per tech day", num1, n_short, True),
             ("service_call_rate", "Service call rate (%)", pct1, p_short, False)]
    fig, axes = plt.subplots(2, 2, figsize=(width, height), dpi=200)
    for ax, (metric, title, fmt, short, good_up) in zip(axes.ravel(), specs):
        vals = m["trends"][metric][market]
        comp = m["trends"][metric]["Company"]
        x = range(len(vals))
        clean = [v for v in vals if v is not None]
        # Rates are comparable to the company on one axis; revenue is not — a $45k
        # market against a $200k company would flatten the market line to nothing.
        compare = metric != "revenue"
        if compare:
            ax.plot(x, comp, color=GRAY, lw=0.9, ls=(0, (3, 2)), zorder=2,
                    label="Company")
        ax.fill_between(x, vals, 0 if metric == "revenue" else (min(clean) * 0.85 if clean else 0),
                        color=BLUE, alpha=0.14, zorder=2)
        ax.plot(x, vals, color=BLUE, lw=1.7, zorder=3, label=market)
        if clean:
            ax.plot([len(vals) - 1], [vals[-1]], "o", ms=3.6, color=NAVY, mec=SURFACE,
                    mew=1.0, zorder=5)
            # Every week labelled, each on the outside of the curve so it clears the
            # line; the last one is called out in full units.
            for i, val in enumerate(vals):
                if val is None:
                    continue
                last = i == len(vals) - 1
                nbrs = [vals[j] for j in (i - 1, i + 1)
                        if 0 <= j < len(vals) and vals[j] is not None]
                above = True if not nbrs else val >= sum(nbrs) / len(nbrs)
                dy = (9 if above else -13) if last else (6 if above else -10)
                ax.annotate(fmt(val) if last else short(val), (i, val),
                            textcoords="offset points",
                            xytext=(0, dy), ha="center",
                            fontsize=6.6 if last else 5.2,
                            color=NAVY if last else MUTED,
                            fontweight="bold" if last else "normal", zorder=6)
            lo = 0 if metric == "revenue" else min(clean) * 0.78
            pool = clean + ([v for v in comp if v is not None] if compare else [])
            span = (max(pool) - lo) or 1
            ax.set_ylim(lo - span * 0.10, max(pool) + span * 0.26)
            ax.set_xlim(-0.8, len(vals) - 0.2)
        ax.set_title(title, fontsize=7.8, color=INK, fontweight="bold", loc="left", pad=3)
        ax.set_yticks([])
        ax.grid(axis="y", color=RULE, lw=0.5, zorder=0)
        ax.set_axisbelow(True)
        for sp in ("top", "right", "left"):
            ax.spines[sp].set_visible(False)
        ax.set_xticks([0, len(vals) - 1])
        ax.set_xticklabels([weeks[0].strftime("%-m/%-d"), weeks[-1].strftime("%-m/%-d")],
                           fontsize=6.4)
    axes[0][0].legend(frameon=False, fontsize=6.2, loc="upper left", handlelength=1.4,
                      borderpad=0, labelspacing=0.15, ncol=2, columnspacing=0.8)
    fig.suptitle("Last 13 operations weeks — market vs company",
                 fontsize=8.8, color=NAVY, fontweight="bold", x=0.012, ha="left", y=0.995)
    fig.tight_layout(pad=0.45, rect=(0, 0, 1, 0.945))
    fig.savefig(path)
    plt.close(fig)


def chart_company_trend(m, path, height=3.30):
    weeks = [dt.date.fromisoformat(w) for w in m["trend_weeks"]]
    vals = [v or 0 for v in m["trend"]["Company"]]
    ma = [None if i < 3 else sum(vals[i-3:i+1]) / 4 for i in range(len(vals))]

    fig, ax = plt.subplots(figsize=(6.05, height), dpi=200)
    x = range(len(weeks))
    bars = ax.bar(x, vals, color=BLUE, width=0.62, zorder=3)
    bars[-1].set_color(NAVY)
    ax.plot(x, ma, color=NAVY, lw=1.6, ls=(0, (4, 2)), zorder=4, label="4-week average")

    for i, v in enumerate(vals):
        ax.annotate(f"{v/1000:.0f}", (i, v), textcoords="offset points", xytext=(0, 3),
                    ha="center", fontsize=6.6,
                    color=NAVY if i == len(vals) - 1 else MUTED,
                    fontweight="bold" if i == len(vals) - 1 else "normal")

    ax.set_xticks(list(x))
    ax.set_xticklabels([w.strftime("%-m/%-d") for w in weeks], fontsize=7)
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"${v/1000:,.0f}k"))
    ax.set_ylim(0, max(vals) * 1.18)
    ax.grid(axis="y", color=RULE, lw=0.6, zorder=0)
    ax.set_axisbelow(True)
    for s in ("top", "right", "left"):
        ax.spines[s].set_visible(False)
    ax.set_ylabel("Net revenue ($000)", fontsize=7.5)
    ax.legend(frameon=False, fontsize=7.5, loc="upper right",
              bbox_to_anchor=(1.0, 1.10), handlelength=2.4)
    ax.set_title("Company net revenue by operations week — label = week start (Thu)",
                 fontsize=9.5, color=NAVY, fontweight="bold", loc="left", pad=10)
    fig.tight_layout(pad=0.6)
    fig.savefig(path)
    plt.close(fig)


def chart_egp(m, path):
    egp = m["egp"]
    rows = [(k, v) for k, v in egp["closed"].items() if k != "Total"]
    rows.sort(key=lambda kv: kv[1]["egp_less_ad_pct"] or -99)
    names = [k for k, _ in rows]
    vals = [(v["egp_less_ad_pct"] or 0) * 100 for _, v in rows]
    total = (egp["closed"]["Total"]["egp_less_ad_pct"] or 0) * 100

    fig, ax = plt.subplots(figsize=(6.05, 4.05), dpi=200)
    colors = [BLUE if v >= 0 else RED for v in vals]
    ax.barh(names, vals, color=colors, height=0.58, zorder=3)
    ax.axvline(0, color=BLACK, lw=1.0, zorder=4)
    ax.axvline(total, color=NAVY, lw=1.3, ls=(0, (4, 2)), zorder=4)
    ax.annotate(f"Company {total:.1f}%", (total, len(names) - 0.30),
                fontsize=7.2, color=NAVY, ha="center", fontweight="bold")

    lo, hi = min(vals + [0]), max(vals + [0])
    span = hi - lo
    for i, v in enumerate(vals):
        off = span * 0.015
        ax.annotate(f"{v:.1f}%", (v + (off if v >= 0 else -off), i), va="center",
                    ha="left" if v >= 0 else "right", fontsize=7.6,
                    color=INK if v >= 0 else RED, fontweight="bold")
    ax.set_xlim(lo - span * 0.14, hi + span * 0.16)
    ax.grid(axis="x", color=RULE, lw=0.6, zorder=0)
    ax.set_axisbelow(True)
    for s in ("top", "right", "left"):
        ax.spines[s].set_visible(False)
    ax.tick_params(axis="y", length=0, labelsize=8.5)
    ax.xaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:.0f}%"))
    ax.set_title(f"External GP less advertising, % of net sales — {egp['meta']['closed_label']} close",
                 fontsize=9.5, color=NAVY, fontweight="bold", loc="left", pad=8)
    fig.tight_layout(pad=0.6)
    fig.savefig(path)
    plt.close(fig)


def chart_small_multiples(m, markets, path, height):
    weeks = [dt.date.fromisoformat(w) for w in m["trend_weeks"]]
    ncol = 2
    nrow = (len(markets) + ncol - 1) // ncol
    fig, axes = plt.subplots(nrow, ncol, figsize=(6.05, height), dpi=200, sharex=True)
    axes = axes.ravel() if hasattr(axes, "ravel") else [axes]

    ordered = sorted(markets, key=lambda k: -(m["markets"][k]["revenue"] or 0))
    peak = max((v or 0) for k in ordered for v in m["trend"][k])
    ymax = peak * 1.30

    for ax, market in zip(axes, ordered):
        vals = [v or 0 for v in m["trend"][market]]
        x = range(len(vals))
        ax.fill_between(x, vals, color=BLUE, alpha=0.16, zorder=2)
        ax.plot(x, vals, color=BLUE, lw=1.6, zorder=3)
        ax.plot([len(vals) - 1], [vals[-1]], "o", ms=4.5, color=NAVY,
                mec=SURFACE, mew=1.2, zorder=4)
        ax.annotate(money_k(vals[-1]), (len(vals) - 1, vals[-1]),
                    textcoords="offset points", xytext=(-2, 6), ha="right",
                    fontsize=7.4, color=NAVY, fontweight="bold")
        wow = m["markets"][market]["revenue_wow"]
        ax.set_title(f"{market}   {signed_pct(wow, 0)} WoW", fontsize=8.2, loc="left",
                     color=INK, fontweight="bold", pad=3)
        ax.set_ylim(0, ymax)
        ax.set_yticks([])
        ax.grid(axis="y", color=RULE, lw=0.5, zorder=0)
        ax.set_axisbelow(True)
        for s in ("top", "right", "left"):
            ax.spines[s].set_visible(False)
        ax.set_xticks([0, len(vals) - 1])
        ax.set_xticklabels([weeks[0].strftime("%-m/%-d"), weeks[-1].strftime("%-m/%-d")],
                           fontsize=6.8)
    for ax in axes[len(ordered):]:
        ax.axis("off")
    scale = f"$0–{money_k(peak)}".replace("$", r"\$")   # avoid matplotlib mathtext
    fig.suptitle(f"Net revenue by operations week — last 13 weeks, common scale {scale}",
                 fontsize=9.5, color=NAVY, fontweight="bold", x=0.012, ha="left", y=0.995)
    fig.tight_layout(pad=0.5, rect=(0, 0, 1, 0.955))
    fig.savefig(path)
    plt.close(fig)


# --- Slide furniture --------------------------------------------------------
def txbox(slide, x, y, w, h, text, size=11, color=INK, bold=False, align=PP_ALIGN.LEFT,
          italic=False, anchor=MSO_ANCHOR.TOP, font=FONT, space_after=2):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
    tf.vertical_anchor = anchor
    lines = text if isinstance(text, list) else [text]
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(space_after)
        runs = line if isinstance(line, list) else [(line, {})]
        for content, style in runs:
            r = p.add_run()
            r.text = content
            r.font.size = Pt(style.get("size", size))
            r.font.bold = style.get("bold", bold)
            r.font.italic = style.get("italic", italic)
            r.font.name = style.get("font", font)
            r.font.color.rgb = rgb(style.get("color", color))
    return box


def rect(slide, x, y, w, h, fill=None, line=None, line_w=0.75):
    from pptx.enum.shapes import MSO_SHAPE
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y),
                                 Inches(w), Inches(h))
    shp.shadow.inherit = False
    if fill:
        shp.fill.solid()
        shp.fill.fore_color.rgb = rgb(fill)
    else:
        shp.fill.background()
    if line:
        shp.line.color.rgb = rgb(line)
        shp.line.width = Pt(line_w)
    else:
        shp.line.fill.background()
    return shp


def slide_frame(prs, title, right_text, sub=None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    txbox(slide, MARGIN, 0.24, 8.6, 0.42, title, size=21, color=NAVY, bold=True)
    txbox(slide, SLIDE_W - MARGIN - 5.0, 0.30, 5.0, 0.30, right_text,
          size=10, color=MUTED, align=PP_ALIGN.RIGHT)
    if sub:
        txbox(slide, MARGIN, 0.64, 12.0, 0.24, sub, size=9, color=MUTED, italic=True)
    rect(slide, MARGIN, 0.92, CONTENT_W, 0.022, fill=NAVY)
    return slide


def kpi_tile(slide, x, y, w, h, label, value, delta_text=None, delta_good=None,
             foot=None, spark=None, value_size=21):
    rect(slide, x, y, w, h, fill=BAND, line=RULE)
    txbox(slide, x + 0.16, y + 0.11, w - 0.32, 0.20, label.upper(), size=8,
          color=MUTED, bold=True)
    txbox(slide, x + 0.16, y + 0.32, w - 0.32, 0.36, value, size=value_size,
          color=NAVY, bold=True)
    if delta_text:
        # A delta with no comparable base reads muted, never as a favourable green.
        if delta_text.lstrip().startswith("—"):
            color, delta_text = MUTED, "no prior-week base"
        else:
            color = GREEN if delta_good else (RED if delta_good is False else MUTED)
        txbox(slide, x + 0.16, y + 0.70, w - 0.32, 0.20, delta_text, size=9.5,
              color=color, bold=True)
    if foot:
        txbox(slide, x + 0.16, y + 0.89, w - 0.32, 0.18, foot, size=7.5, color=MUTED)
    if spark:
        slide.shapes.add_picture(spark, Inches(x + 0.14), Inches(y + h - 0.50),
                                 width=Inches(w - 0.28))


def heat_fill(value, values, good_up):
    """Rank a value against its peers and return a tint. Fewer than three peers gets
    best/worst only — a two-technician market has no meaningful middle."""
    vals = sorted(v for v in values if v is not None)
    if value is None or len(vals) < 2:
        return None
    lo, hi = vals[0], vals[-1]
    if hi - lo < 1e-9:
        return None
    t = (value - lo) / (hi - lo)
    if not good_up:
        t = 1 - t
    if len(vals) == 2:
        return HEAT[4] if t > 0.5 else HEAT[0]
    idx = 0 if t < 0.2 else 1 if t < 0.4 else 2 if t < 0.6 else 3 if t < 0.8 else 4
    return HEAT[idx]


MARKER = {"down": "▼", "up": "▲"}
MARKER_COLOR = {"down": RED, "up": GREEN}


def flag_reasons(market, tech_block):
    """Plain-English reason per marked technician, so a marker never needs
    interpreting. Returns [(name, marker, reason)] in the table's own order."""
    data = tech_block["markets"].get(market)
    if not data:
        return []
    out = []
    for name, v in sorted(data["technicians"].items(),
                          key=lambda kv: -(kv[1]["revenue"] or 0)):
        self_ = v.get("self") or {}
        flags, marker = self_.get("flags") or [], self_.get("marker")
        if not flags or not marker:
            continue
        # Only label each movement with its own arrow when a technician has moved both
        # ways at once; otherwise the name's marker already says the direction.
        mixed = len({f["direction"] for f in flags}) > 1
        parts = []
        for f in flags:
            arrow = f"{MARKER[f['direction']]} " if mixed else ""
            if f["metric"] == "avg_ticket":
                parts.append(f"{arrow}ticket {money(f['now'])} vs own "
                             f"{money(f['before'])} ({signed_pct(f['delta'], 0)})")
            elif f["metric"] == "service_call_rate":
                parts.append(f"{arrow}service call {pct(f['now'], 1)} vs own "
                             f"{pct(f['before'], 1)} ({f['delta']*100:+.1f}pp)")
        out.append((name, marker, "; ".join(parts)))
    return out


def pto_note(market, tech_block):
    """Who in this market was on paid time off during the leaderboard window, and for
    how long. Volume columns move with days worked, so someone out for a week reads low
    on jobs without anything being wrong. Reported, never netted off: the leaderboard
    still shows what happened, this says who was there for it.

    Returns one line, or "" when nobody in the market took time off."""
    data = tech_block["markets"].get(market)
    if not data or not tech_block.get("has_pto"):
        return ""
    out = [(n, v["pto_hours"]) for n, v in data["technicians"].items()
           if v.get("pto_hours")]
    if not out:
        return ""
    out.sort(key=lambda kv: -kv[1])
    return "PTO in window — " + ", ".join(f"{n} {h:,.0f}h" for n, h in out) + "."


def leaderboard_table(slide, x, y, w, h, market, tech_block, size=7.6):
    """The technician heat map, rebuilt from the published feed."""
    data = tech_block["markets"].get(market)
    if not data or not data["technicians"]:
        txbox(slide, x, y, w, 0.3, "No technician activity in this window.",
              size=9, color=MUTED, italic=True)
        return
    techs = sorted(data["technicians"].items(), key=lambda kv: -(kv[1]["revenue"] or 0))
    cols = LEADERBOARD
    columns = {key: [v.get(key) for _, v in techs] for key, *_ in cols}

    n_rows, n_cols = len(techs) + 2, len(cols) + 1
    row_h = min(0.28, h / n_rows)
    shape = slide.shapes.add_table(n_rows, n_cols, Inches(x), Inches(y), Inches(w),
                                   Inches(row_h * n_rows))
    tbl = shape.table
    tbl.first_row = False
    tbl.horz_banding = False
    widths = [1.24] + [0.42, 0.74, 0.44, 0.62, 0.64, 0.58, 0.66, 0.50, 0.64, 0.54, 0.56]
    scale = w / sum(widths)
    for i, cw in enumerate(widths):
        tbl.columns[i].width = Inches(cw * scale)
    for r in range(n_rows):
        tbl.rows[r].height = Inches(row_h)

    def put(cell, text, *, bold=False, color=INK, fill=None, align=PP_ALIGN.RIGHT,
            sz=size):
        cell.margin_left = cell.margin_right = Inches(0.04)
        cell.margin_top = cell.margin_bottom = Inches(0.01)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        if fill:
            cell.fill.solid()
            cell.fill.fore_color.rgb = rgb(fill)
        else:
            cell.fill.background()
        p = cell.text_frame.paragraphs[0]
        cell.text_frame.word_wrap = False
        p.alignment = align
        run = p.add_run()
        run.text = text
        run.font.size = Pt(sz)
        run.font.bold = bold
        run.font.name = FONT
        run.font.color.rgb = rgb(color)

    put(tbl.cell(0, 0), "TECHNICIAN", bold=True, color="#FFFFFF", fill=NAVY,
        align=PP_ALIGN.LEFT, sz=6.8)
    for c, (_, label, *_rest) in enumerate(cols, start=1):
        put(tbl.cell(0, c), label, bold=True, color="#FFFFFF", fill=NAVY, sz=6.8)

    for r, (name, v) in enumerate(techs, start=1):
        # An arrow marks a technician who has stepped away from their OWN prior
        # baseline, either way — a bad patch or a recovery, both worth naming.
        marker = (v.get("self") or {}).get("marker")
        put(tbl.cell(r, 0), (MARKER[marker] + " " if marker else "") + name, bold=True,
            color=MARKER_COLOR.get(marker, INK), align=PP_ALIGN.LEFT, sz=size)
        for c, (key, _label, fmt, shade, good_up) in enumerate(cols, start=1):
            fill = heat_fill(v.get(key), columns[key], good_up) if shade else None
            put(tbl.cell(r, c), FORMATTERS[fmt](v.get(key)), fill=fill, sz=size)

    tot = data["total"]
    put(tbl.cell(n_rows - 1, 0), f"{market} overall", bold=True, fill="#EDF1F5",
        align=PP_ALIGN.LEFT, sz=size)
    for c, (key, _label, fmt, *_rest) in enumerate(cols, start=1):
        put(tbl.cell(n_rows - 1, c), FORMATTERS[fmt](tot.get(key)), bold=True,
            fill="#EDF1F5", sz=size)


def add_table(slide, x, y, w, h, header, rows, widths, aligns, header_fill=NAVY,
              size=8.5, header_size=8, row_h=None, total_row=False):
    n_rows, n_cols = len(rows) + 1, len(header)
    shape = slide.shapes.add_table(n_rows, n_cols, Inches(x), Inches(y),
                                   Inches(w), Inches(h))
    tbl = shape.table
    tbl.first_row = False
    tbl.horz_banding = False
    scale = w / sum(widths)
    for i, cw in enumerate(widths):
        tbl.columns[i].width = Inches(cw * scale)
    rh = row_h or (h / n_rows)
    for r in range(n_rows):
        tbl.rows[r].height = Inches(rh)

    def style(cell, text, *, bold=False, color=INK, fill=None, align=PP_ALIGN.RIGHT,
              sz=size):
        cell.margin_left = cell.margin_right = Inches(0.05)
        cell.margin_top = cell.margin_bottom = Inches(0.01)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        if fill:
            cell.fill.solid()
            cell.fill.fore_color.rgb = rgb(fill)
        else:
            cell.fill.background()
        tf = cell.text_frame
        tf.word_wrap = False
        p = tf.paragraphs[0]
        p.alignment = align
        r = p.add_run()
        r.text = text
        r.font.size = Pt(sz)
        r.font.bold = bold
        r.font.name = FONT
        r.font.color.rgb = rgb(color)

    for c, label in enumerate(header):
        style(tbl.cell(0, c), label, bold=True, color="#FFFFFF", fill=header_fill,
              align=aligns[c], sz=header_size)

    for r, row in enumerate(rows, start=1):
        is_total = total_row and r == n_rows - 1
        base_fill = "#EDF1F5" if is_total else (BAND if r % 2 == 0 else None)
        for c, cellspec in enumerate(row):
            text, color = (cellspec if isinstance(cellspec, tuple) else (cellspec, INK))
            style(tbl.cell(r, c), text, bold=is_total or c == 0, color=color,
                  fill=base_fill, align=aligns[c])
    return tbl


def owner_box(slide, x, y, w, h, title, prompt):
    rect(slide, x, y, w, h, fill=SURFACE, line=RULE)
    rect(slide, x, y, w, 0.30, fill=NAVY)
    txbox(slide, x + 0.11, y + 0.065, w - 0.22, 0.20, title.upper(), size=8.5,
          color="#FFFFFF", bold=True)
    txbox(slide, x + 0.11, y + 0.40, w - 0.22, h - 0.50, prompt, size=8,
          color="#AEB6BF", italic=True)


def bullet_chars(b):
    if isinstance(b, list):
        return sum(len(t) for t, _ in b)
    return len(b)


def band_height(bullets, width_in, size=9.5, pad=0.46):
    """Height needed for a callout band, counting wrapped lines."""
    chars_per_line = max(40, int(width_in * 12.8))  # ~12.8 chars/inch at 9.5pt Calibri
    lines = sum(max(1, -(-bullet_chars(b) // chars_per_line)) for b in bullets)
    return pad + lines * 0.185


def callout_band(slide, x, y, w, h, heading, bullets):
    rect(slide, x, y, w, h, fill=BAND, line=RULE)
    txbox(slide, x + 0.16, y + 0.10, w - 0.32, 0.20, heading.upper(), size=8,
          color=MUTED, bold=True)
    lines = []
    for b in bullets:
        if isinstance(b, list):
            lines.append([("•  ", {"color": BLUE, "bold": True})] + b)
        else:
            lines.append([("•  ", {"color": BLUE, "bold": True}), (b, {})])
    txbox(slide, x + 0.16, y + 0.33, w - 0.32, h - 0.43, lines, size=9.5,
          color=INK, space_after=3)


# --- Narrative --------------------------------------------------------------
def describe_outlier(o):
    d = "up" if o["direction"] == "up" else "down"
    if o["metric"] == "revenue":
        detail = f"{money(o['value'])} vs a {money(o['baseline'])} 8-week norm"
    elif o["metric"] == "avg_ticket":
        detail = f"{money(o['value'])} vs {money(o['baseline'])}"
    elif o["metric"] == "service_call_rate":
        detail = f"{pct(o['value'])} vs {pct(o['baseline'])}"
    else:
        detail = f"{num(o['value'])} vs {num(o['baseline'])}"
    return [(f"{o['market']}: ", {"bold": True}),
            (f"{o['label'].lower()} {d} {abs(o['delta_pct'])*100:.0f}% — {detail}.",
             {"color": GREEN if o["favourable"] else RED if abs(o["z"]) > 2.5 else INK})]


def company_callouts(m):
    out = []
    c, b = m["company"], m["bridge"]["Company"]
    if m["holidays"]:
        names = " and ".join(h["name"] for h in m["holidays"])
        out.append([(f"{names} falls in this week. ", {"bold": True}),
                    ("Compare rates and per-tech-day figures, not totals.", {})])
    if b:
        share = b["capacity_share"]
        lead = "capacity" if share and share >= 0.5 else "productivity"
        out.append([("Revenue bridge: ", {"bold": True}),
                    (f"{signed_pct(c['revenue_wow'], 1)} week over week splits "
                     f"{money(b['capacity_effect'])} capacity "
                     f"({b['tech_days_delta']:+.0f} tech days) and "
                     f"{money(b['productivity_effect'])} productivity "
                     f"({money(b['rev_per_tech_day'])}/day vs "
                     f"{money(b['rev_per_tech_day_prior'])}). "
                     f"{lead.title()} is the larger driver.", {})])
    ms = m["monthly"]["series"]
    cur, prev = ms[-1], ms[-2]
    out.append([("Month pace: ", {"bold": True}),
                (f"{money(cur['total'])} through day {m['monthly']['days_elapsed']} of "
                 f"{m['monthly']['days_in_month']} paces to {money(cur['pace'])}, "
                 f"{signed_pct(pct_delta(cur['pace'], prev['total']), 1)} against "
                 f"{prev['label']} at {money(prev['total'])}. "
                 f"Straight-line pace — no LE loaded.", {})])
    for o in m["outliers"]["company"][:2]:
        out.append(describe_outlier(o))
    if m.get("data_quality", {}).get("feed_gap_flag"):
        gap = m["data_quality"]["feed_revenue_gap"]
        out.append([("Data note: ", {"bold": True}),
                    (f"the technician feed reports {signed_pct(gap, 1)} more revenue than "
                     "the channel feed this week (normally under 0.5%). Revenue figures "
                     "here use the channel feed throughout; late invoice attribution is "
                     "the likely cause and it usually settles within a week.", {})])
    return out[:6]


def region_observations(m, region):
    markets = m["regions"][region]["markets"]
    out = []
    rev = sum(m["markets"][k]["revenue"] or 0 for k in markets)
    prior = sum(m["markets"][k]["revenue_prior"] or 0 for k in markets)
    delta = (rev - prior) / prior if prior else None
    days = sum((m["bridge"][k] or {}).get("tech_days", 0) for k in markets)
    days_prior = sum((m["bridge"][k] or {}).get("tech_days_prior", 0) for k in markets)
    out.append([(f"Region revenue {money(rev)}, {signed_pct(delta, 1)} week over week", {"bold": True}),
                (f"; {days:.0f} tech days worked vs {days_prior:.0f} prior week "
                 f"({money(rev/days) if days else '—'} per tech day).", {})])
    for o in m["outliers"][region][:3]:
        out.append(describe_outlier(o))
    if len(out) == 1:
        out.append("No market cleared the outlier threshold this week — all metrics "
                   "within 1.5 standard deviations of their 8-week norm.")
    return out[:4]


# --- Slides -----------------------------------------------------------------
def exec_summary(prs, m):
    wk = m["ops_week"]
    stamp = ("Data complete through " if m["complete"] else "PARTIAL — data through ") + \
        fmt_date(m["feed_last_day"]["ds01"])
    slide = slide_frame(prs, "Executive Performance Summary",
                        f"Operations week {wk['label']}  ·  {stamp}",
                        "All markets. Net revenue excludes the Fleet channel. "
                        "Week runs Thursday–Wednesday.")
    c = m["company"]
    tiles = [
        ("Net revenue", money(c["revenue"]), signed_pct(c["revenue_wow"], 1) + " WoW",
         (c["revenue_wow"] or 0) >= 0, f"8-week norm {money(c['revenue_baseline'])}",
         "revenue", True),
        ("Average ticket", money(c["avg_ticket"]), signed_pct(c["avg_ticket_wow"], 1) + " WoW",
         (c["avg_ticket_wow"] or 0) >= 0, f"8-week norm {money(c['avg_ticket_baseline'])}",
         "avg_ticket", True),
        ("Rev-producing jobs / tech day", num(c["jobs_per_tech_day"]),
         signed_pct(c["jobs_per_tech_day_wow"], 1) + " WoW",
         (c["jobs_per_tech_day_wow"] or 0) >= 0,
         f"{c['active_techs']} techs · {c['tech_days']:.0f} tech days",
         "jobs_per_tech_day", True),
        ("Service call rate", pct(c["service_call_rate"]),
         signed_pct(c["service_call_rate_wow"], 1) + " WoW",
         (c["service_call_rate_wow"] or 0) <= 0,
         f"8-week norm {pct(c['service_call_rate_baseline'])}",
         "service_call_rate", False),
    ]
    # Technician efficiency = invoiced hours ÷ clocked hours. Clocked hours land Friday
    # afternoon, after this deck ships, so the tile states that rather than guessing.
    tech_eff = (m.get("technicians") or {}).get("markets", {})
    # Name the week payroll still owes, so a dash reads as a known hand-off lag rather
    # than something broken. Partial cover is never shown: see compute.technician_table.
    gap = (m.get("technicians") or {}).get("clocked_weeks_missing") or []
    if gap:
        since = dt.date.fromisoformat(min(gap))
        foot = f"Awaiting payroll for week of {since.strftime('%b %-d')}"
    else:
        foot = "Clocked hours not yet in the feed"
    eff_tile = ("Technician efficiency", "—", None, None, foot, None, True)
    if (m.get("technicians") or {}).get("has_clocked_hours"):
        inv = sum(d["total"]["invoiced_hours"] or 0 for d in tech_eff.values())
        clk = sum(d["total"].get("clocked_hours") or 0 for d in tech_eff.values())
        eff_tile = ("Technician efficiency", pct(inv / clk if clk else None, 0), None, None,
                    f"{inv:,.0f} invoiced ÷ {clk:,.0f} clocked hours", None, True)
    tiles.append(eff_tile)

    tw = (CONTENT_W - 4 * 0.14) / 5
    for i, (label, value, delta, good, foot, metric, good_up) in enumerate(tiles):
        spark = None
        if metric:
            spark = os.path.join(m["_out"], f"spark_{metric}.png")
            chart_sparkline(m["trends"][metric]["Company"], spark, good_up=good_up,
                            width=tw - 0.28)
        kpi_tile(slide, MARGIN + i * (tw + 0.14), 1.02, tw, 1.62, label, value, delta,
                 good, foot, spark=spark, value_size=19)

    bullets = company_callouts(m)
    band_h = band_height(bullets, CONTENT_W - 0.32)
    band_y = SLIDE_H - 0.32 - band_h
    top, body_h = 2.82, band_y - 0.14 - 2.82

    chart = os.path.join(m["_out"], "chart_month_pace.png")
    chart_month_pace(m, chart, height=body_h)
    slide.shapes.add_picture(chart, Inches(MARGIN), Inches(top), width=Inches(6.05))

    # Month pace per market: the in-flight month extrapolated from elapsed days,
    # measured against the last closed month.
    ms = m["monthly"]["series"]
    cur, prev = ms[-1], ms[-2]

    def pace_vs_lm(market):
        return pct_delta((cur.get("market_pace") or {}).get(market),
                         prev["markets"].get(market))

    rows = []
    ordered = sorted(m["markets"].values(), key=lambda v: -(v["revenue"] or 0))
    for v in ordered:
        wow, twow = v["revenue_wow"], v["avg_ticket_wow"]
        pv = pace_vs_lm(v["market"])
        rows.append([
            v["market"], money(v["revenue"]),
            (signed_pct(wow, 0), GREEN if (wow or 0) >= 0 else RED),
            money(v["avg_ticket"]),
            (signed_pct(twow, 0), GREEN if (twow or 0) >= 0 else RED),
            num(v["jobs_per_tech_day"]), pct(v["service_call_rate"], 1),
            (signed_pct(pv, 0), GREEN if (pv or 0) >= 0 else RED),
        ])
    cw = m["company"]
    tot_pace = pct_delta(cur["pace"], prev["total"])
    rows.append(["Total", money(cw["revenue"]),
                 (signed_pct(cw["revenue_wow"], 0), GREEN if (cw["revenue_wow"] or 0) >= 0 else RED),
                 money(cw["avg_ticket"]),
                 (signed_pct(cw["avg_ticket_wow"], 0), GREEN if (cw["avg_ticket_wow"] or 0) >= 0 else RED),
                 num(cw["jobs_per_tech_day"]), pct(cw["service_call_rate"], 1),
                 (signed_pct(tot_pace, 0), GREEN if (tot_pace or 0) >= 0 else RED)])
    tbl_h = body_h - 0.24
    txbox(slide, 6.73, top, 6.18, 0.20,
          f"MARKET SCORECARD — WEEK {wk['label'].upper()}, SORTED BY REVENUE",
          size=8, color=NAVY, bold=True)
    add_table(slide, 6.73, top + 0.24, 6.18, tbl_h,
              ["Market", "Revenue", "WoW", "Ticket", "WoW", "Rev jobs/day", "Svc call %",
               f"{cur['label']} pace vs {prev['label']}"],
              rows, [1.06, 0.90, 0.60, 0.76, 0.60, 0.68, 0.72, 0.98],
              [PP_ALIGN.LEFT] + [PP_ALIGN.RIGHT] * 7, total_row=True,
              row_h=tbl_h / (len(rows) + 1), size=8.0, header_size=6.6)

    callout_band(slide, MARGIN, band_y, CONTENT_W, band_h,
                 "What changed this week — auto-generated", bullets)
    return slide


def profitability(prs, m):
    egp = m["egp"]
    meta = egp["meta"]
    slide = slide_frame(
        prs, "Profitability & Unit Economics",
        f"{meta['closed_label']} close  ·  {meta['open_label']} MTD day "
        f"{meta['days_elapsed']}/{meta['days_in_month']}  ·  model snapshot "
        f"{fmt_date(meta['pulled'])}",
        "Source: EGP (ex Ad) Optimization Model — NCM dashboard tabs. "
        "External GP less advertising is the market-health standard; every market "
        "must be positive.")

    chart = os.path.join(m["_out"], "chart_egp.png")
    chart_egp(m, chart)
    slide.shapes.add_picture(chart, Inches(MARGIN), Inches(1.02), width=Inches(6.05))

    closed, pacing = egp["closed"], egp["open_pacing"]
    labor = m.get("labor")
    lab_months = labor["months"] if labor else []
    lab_labels = labor["labels"] if labor else []

    def lab(market, month):
        if not labor:
            return None
        return (labor["markets"].get(market, {}).get(month) or {}).get("pct")

    def lab_cell(market, month):
        v = lab(market, month)
        if v is None:
            return ("—", MUTED)
        # Labor is a cost: rising is bad. Colour only the last month, against the first,
        # so the eye reads direction rather than ten independent reds.
        if month != lab_months[-1] or len(lab_months) < 2:
            return (pct(v, 1), INK)
        first = lab(market, lab_months[0])
        if first is None:
            return (pct(v, 1), INK)
        return (pct(v, 1), RED if v > first else GREEN)

    rows = []
    ordered = sorted([k for k in closed if k != "Total"],
                     key=lambda k: -(closed[k]["net_sales"] or 0))
    for k in ordered:
        c, p = closed[k], (pacing or {}).get(k, {})
        ela = c["egp_less_ad"]
        rows.append([
            k, money_k(c["net_sales"]), pct(c["gp_pct"], 0),
            (money_k(ela), GREEN if (ela or 0) >= 0 else RED),
            (pct(c["egp_less_ad_pct"], 1), GREEN if (c["egp_less_ad_pct"] or 0) >= 0 else RED),
            (pct(p.get("egp_less_ad_pct"), 1),
             GREEN if (p.get("egp_less_ad_pct") or 0) >= 0 else RED),
        ] + [lab_cell(k, mo) for mo in lab_months])
    t, tp = closed["Total"], (pacing or {}).get("Total", {})
    rows.append(["Total", money_k(t["net_sales"]), pct(t["gp_pct"], 0),
                 (money_k(t["egp_less_ad"]), GREEN), (pct(t["egp_less_ad_pct"], 1), GREEN),
                 (pct(tp.get("egp_less_ad_pct"), 1),
                  GREEN if (tp.get("egp_less_ad_pct") or 0) >= 0 else RED)]
                + [lab_cell("Total", mo) for mo in lab_months])

    txbox(slide, 6.73, 1.02, 6.18, 0.20,
          f"MARKET PROFITABILITY — {meta['closed_label'].upper()} CLOSE", size=8,
          color=NAVY, bold=True)
    header = ["Market", "Net sales", "GP %", "EGP less ad", "EGP ex ad %",
              f"{meta['open_label'][:3]} pace %"] + [f"Labor % {l}" for l in lab_labels]
    widths = [0.98, 0.78, 0.50, 0.82, 0.76, 0.76] + [0.60] * len(lab_months)
    add_table(slide, 6.73, 1.26, 6.18, 3.72, header, rows, widths,
              [PP_ALIGN.LEFT] + [PP_ALIGN.RIGHT] * (len(header) - 1), total_row=True,
              row_h=3.72 / 12, size=7.8, header_size=6.4)
    txbox(slide, 6.73, 5.02, 6.18, 0.34,
          "Labor % is the P&L Cost of Sales labor line over net revenue. Salaried "
          "ops-admin pay sits in Market G&A, not here. Last month coloured against the "
          "first — red means labor is taking a larger share of sales.",
          size=6.8, color=MUTED, italic=True)

    fails = [k for k in ordered if (closed[k]["egp_less_ad"] or 0) <= 0]
    thin = [k for k in ordered
            if 0 < (closed[k]["egp_less_ad_pct"] or 0) < 0.05]
    pace_fails = [k for k in ordered
                  if pacing and (pacing.get(k, {}).get("egp_less_ad") or 0) <= 0]
    bullets = []
    bullets.append([("Standard: ", {"bold": True}),
                    (f"{len(ordered) - len(fails)} of {len(ordered)} markets cleared a "
                     f"positive EGP less advertising in {meta['closed_label']}. "
                     f"Company {pct(t['egp_less_ad_pct'], 1)} of net sales "
                     f"({money(t['egp_less_ad'])}).", {})])
    if fails:
        verb = "does not" if len(fails) == 1 else "do not"
        subject = "This market" if len(fails) == 1 else "These markets"
        bullets.append([("Below standard: ", {"bold": True, "color": RED}),
                        (", ".join(f"{k} ({money(closed[k]['egp_less_ad'])})" for k in fails)
                         + f". {subject} {verb} cover its own advertising."
                         if len(fails) == 1 else
                         ", ".join(f"{k} ({money(closed[k]['egp_less_ad'])})" for k in fails)
                         + f". {subject} {verb} cover their own advertising.",
                         {"color": RED})])
    if thin:
        bullets.append([("Thin: ", {"bold": True}),
                        (", ".join(f"{k} ({pct(closed[k]['egp_less_ad_pct'], 1)})" for k in thin)
                         + " — positive but under 5% of net sales.", {})])
    if pace_fails:
        bullets.append([(f"{meta['open_label']} pacing negative: ", {"bold": True}),
                        (", ".join(pace_fails) + ". Pacing is the NCM model's own "
                         "full-month extrapolation from MTD, not a close.", {})])
    if labor and len(lab_months) >= 2:
        first, last = lab_months[0], lab_months[-1]
        scored = [(k, lab(k, last), lab(k, first)) for k in ordered
                  if lab(k, last) is not None and lab(k, first) is not None]
        if scored:
            worst = max(scored, key=lambda s: s[1])
            risers = sorted((s for s in scored if s[1] - s[2] > 0.02),
                            key=lambda s: -(s[1] - s[2]))[:3]
            tot_l, tot_f = lab("Total", last), lab("Total", first)
            line = [("Labor: ", {"bold": True}),
                    (f"company {pct(tot_l, 1)} of sales in {lab_labels[-1]} vs "
                     f"{pct(tot_f, 1)} in {lab_labels[0]}. Highest is {worst[0]} at "
                     f"{pct(worst[1], 1)}.", {})]
            if risers:
                line.append((" Rising fastest: " + ", ".join(
                    f"{k} {pct(f, 1)}→{pct(l, 1)}" for k, l, f in risers) + ".",
                    {"color": RED}))
            bullets.append(line)
    bullets.append([("Read with care: ", {"bold": True}),
                    (f"{meta['closed_label']} is final. {meta['open_label']} pacing is an "
                     f"estimate from {meta['days_elapsed']} of {meta['days_in_month']} days "
                     "and will move. EBITDA is not in this model — it lands with the "
                     "monthly P&L.", {})])
    band_h = max(1.55, band_height(bullets, CONTENT_W - 0.32))
    callout_band(slide, MARGIN, SLIDE_H - 0.32 - band_h, CONTENT_W, band_h,
                 "Read-out — auto-generated", bullets)
    return slide


def region_slide(prs, m, region):
    cfg = m["regions"][region]
    owner, markets = cfg["owner"], cfg["markets"]
    wk = m["ops_week"]
    slide = slide_frame(prs, f"{region} Operations Review — {owner}",
                        f"Operations week {wk['label']}",
                        f"{', '.join(markets)}. Roll-up view — the action plan for each "
                        f"market sits on its own slide behind this one.")

    h = 2.62 if len(markets) <= 4 else 2.86
    chart = os.path.join(m["_out"], f"chart_{region.lower()}.png")
    chart_small_multiples(m, markets, chart, h)
    slide.shapes.add_picture(chart, Inches(MARGIN), Inches(1.02), width=Inches(6.05))

    rows = []
    ordered = sorted(markets, key=lambda k: -(m["markets"][k]["revenue"] or 0))
    agg = {"revenue": 0.0, "revenue_prior": 0.0, "jobs_completed": 0, "tech_days": 0.0,
           "tech_jobs": 0, "service_calls": 0, "tech_revenue": 0.0, "booked_revenue": 0.0}
    for k in ordered:
        v = m["markets"][k]
        wow, base = v["revenue_wow"], v["revenue_vs_baseline"]
        rows.append([
            k, money(v["revenue"]),
            (signed_pct(wow, 0), GREEN if (wow or 0) >= 0 else RED),
            (signed_pct(base, 0), GREEN if (base or 0) >= 0 else RED),
            money(v["avg_ticket"]), pct(v["ticket_completion"], 1),
            num(v["jobs_per_tech_day"]), pct(v["service_call_rate"], 1),
        ])
        for key in agg:
            agg[key] += v.get(key) or 0
    tot_wow = (agg["revenue"] - agg["revenue_prior"]) / agg["revenue_prior"] \
        if agg["revenue_prior"] else None
    dispatches = agg["tech_jobs"] + agg["service_calls"]
    rows.append(["Region total", money(agg["revenue"]),
                 (signed_pct(tot_wow, 0), GREEN if (tot_wow or 0) >= 0 else RED), "—",
                 money(agg["revenue"] / agg["jobs_completed"]) if agg["jobs_completed"] else "—",
                 pct(agg["tech_revenue"] / agg["booked_revenue"], 1) if agg["booked_revenue"] else "—",
                 num(agg["tech_jobs"] / agg["tech_days"]) if agg["tech_days"] else "—",
                 pct(agg["service_calls"] / dispatches, 1) if dispatches else "—"])

    n = len(rows) + 1
    txbox(slide, 6.73, 1.02, 6.18, 0.20, f"{region.upper()} MARKET DETAIL", size=8,
          color=NAVY, bold=True)
    add_table(slide, 6.73, 1.26, 6.18, min(0.30 * n, 2.62),
              ["Market", "Revenue", "WoW", "vs 8wk", "Ticket", "Ticket compl.",
               "Rev jobs/day", "Svc call %"],
              rows, [1.10, 0.90, 0.62, 0.62, 0.74, 0.86, 0.68, 0.76],
              [PP_ALIGN.LEFT] + [PP_ALIGN.RIGHT] * 7, total_row=True,
              row_h=min(0.30, 2.62 / n), size=8.0, header_size=7.0)

    obs = region_observations(m, region)
    band_y = 1.02 + h + 0.08                      # directly under the small multiples
    band_h = max(0.92, band_height(obs, CONTENT_W - 0.32))
    callout_band(slide, MARGIN, band_y, CONTENT_W, band_h,
                 "What the data says — auto-generated", obs)

    # Trend band: the three rate metrics over 13 weeks, region against its markets.
    specs = [("avg_ticket", "Average ticket", money0, True),
             ("jobs_per_tech_day", "Jobs per tech day", num1, True),
             ("service_call_rate", "Service call rate", pct1, False)]
    cw_ = (CONTENT_W - 2 * 0.14) / 3
    trend_y = band_y + band_h + 0.10
    trend_h = SLIDE_H - 0.32 - trend_y
    for i, (metric, title, fmt, good_up) in enumerate(specs):
        p = os.path.join(m["_out"], f"trend_{region.lower()}_{metric}.png")
        chart_region_trend(m, region, metric, title, p, width=cw_, height=trend_h,
                           fmt=fmt, good_up=good_up)
        slide.shapes.add_picture(p, Inches(MARGIN + i * (cw_ + 0.14)), Inches(trend_y),
                                 width=Inches(cw_))
    return slide


def market_slide(prs, m, market):
    """One market, one action plan. Trends in the company's language on the left,
    the technician heat map on the right, the manager's plan underneath."""
    region = next(r for r, cfg in m["regions"].items() if market in cfg["markets"])
    owner = m["regions"][region]["owner"]
    wk = m["ops_week"]
    v = m["markets"][market]
    tb = m["technicians"]

    slide = slide_frame(
        prs, f"{market} — {region} · {owner}",
        f"Operations week {wk['label']}",
        f"Technician table covers the trailing {tb['weeks']} weeks "
        f"({fmt_date(tb['start'])}–{fmt_date(tb['end'])}) — a single week is too thin "
        f"for a fair per-technician read.")

    ms = m["monthly"]["series"]
    pv = pct_delta((ms[-1].get("market_pace") or {}).get(market),
                   ms[-2]["markets"].get(market))

    tiles = [
        ("Net revenue", money(v["revenue"]), signed_pct(v["revenue_wow"], 1) + " WoW",
         (v["revenue_wow"] or 0) >= 0,
         f"{ms[-1]['label']} pacing {signed_pct(pv, 1)} vs {ms[-2]['label']} "
         f"(day {m['monthly']['days_elapsed']})"),
        ("Average ticket", money(v["avg_ticket"]),
         signed_pct(v["avg_ticket_wow"], 1) + " WoW",
         (v["avg_ticket_wow"] or 0) >= 0,
         f"Ticket completion {pct(v['ticket_completion'], 1)} · "
         f"company {money(m['company']['avg_ticket'])}"),
        ("Rev-producing jobs / tech day", num(v["jobs_per_tech_day"]),
         signed_pct(v["jobs_per_tech_day_wow"], 1) + " WoW",
         (v["jobs_per_tech_day_wow"] or 0) >= 0,
         f"{v['active_techs']} techs · {v['tech_days']:.0f} tech days"),
        ("Service call rate", pct(v["service_call_rate"]),
         signed_pct(v["service_call_rate_wow"], 1) + " WoW",
         (v["service_call_rate_wow"] or 0) <= 0,
         f"Company {pct(m['company']['service_call_rate'])}"),
    ]
    tw = (CONTENT_W - 3 * 0.14) / 4
    for i, (label, value, delta, good, foot) in enumerate(tiles):
        kpi_tile(slide, MARGIN + i * (tw + 0.14), 1.02, tw, 1.00, label, value, delta,
                 good, foot, value_size=19)

    panel = os.path.join(m["_out"], f"panel_{market.lower().replace(' ', '_')}.png")
    chart_market_panel(m, market, panel, width=5.26, height=2.62)
    slide.shapes.add_picture(panel, Inches(MARGIN), Inches(2.16), width=Inches(5.26))

    txbox(slide, 5.86, 2.16, 7.05, 0.20,
          f"TECHNICIAN LEADERBOARD — TRAILING {tb['weeks']} WEEKS", size=8,
          color=NAVY, bold=True)
    leaderboard_table(slide, 5.86, 2.40, 7.05, 2.28, market, tb)

    # Why each ▼ fired, in words. A marker nobody can explain gets ignored or argued
    # with; one that names the number and the technician's own norm gets acted on.
    reasons = flag_reasons(market, tb)
    note_y = 4.72
    if reasons:
        head = [[(f"▲ ▼ = moved against their own prior {tb['baseline_weeks']} weeks, "
                  f"not a peer ranking:", {"color": MUTED, "italic": True})]]
        lines = head + [[(MARKER[mk] + " ", {"color": MARKER_COLOR[mk], "bold": True}),
                         (f"{name} — ", {"bold": True, "color": MARKER_COLOR[mk]}),
                         (why, {})] for name, mk, why in reasons]
        h = 0.145 * len(lines) + 0.05
        txbox(slide, 5.86, note_y, 7.05, h, lines, size=6.9, color=INK, space_after=0)
        note_y += h + 0.02
    # Time off sits with the reasons rather than in the table: it explains a thin
    # volume column, it is not a metric anyone is ranked on.
    pto = pto_note(market, tb)
    if pto:
        txbox(slide, 5.86, note_y, 7.05, 0.16,
              [[("PTO in window — ", {"bold": True, "color": MUTED}),
                (pto.split("— ", 1)[1], {"color": MUTED})]],
              size=6.9, color=MUTED, space_after=0)
        note_y += 0.18
    txbox(slide, 5.86, note_y, 7.05, 0.18,
          "Shading ranks within this market. Volume columns and booked ticket are "
          "unshaded — booked ticket is the work dispatched, not the result.",
          size=6.4, color=MUTED, italic=True)
    box_y = max(5.12, note_y + 0.24)

    boxes = [
        ("Key takeaways", f"{owner}: what is happening in {market}, in two or three lines."),
        ("Current priorities", f"The one or two things being worked in {market} this "
                               "week, and why they come first here."),
        ("Actions underway", "Specific action, owner, and the date it lands."),
        ("Support / decisions needed", "What is blocked, who needs to unblock it, and "
                                       "by when."),
    ]
    bw = (CONTENT_W - 3 * 0.14) / 4
    box_h = SLIDE_H - 0.32 - box_y
    for i, (title, prompt) in enumerate(boxes):
        owner_box(slide, MARGIN + i * (bw + 0.14), box_y, bw, box_h, title, prompt)
    return slide


def title_note(prs, m):
    """Trailing method note — keeps definitions out of the operating slides."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    txbox(slide, MARGIN, 0.24, 9.0, 0.42, "Method & definitions", size=21,
          color=NAVY, bold=True)
    rect(slide, MARGIN, 0.92, CONTENT_W, 0.022, fill=NAVY)
    lines = [
        [("Operations week. ", {"bold": True}),
         ("Thursday through Wednesday. The deck is generated Thursday night for the "
          "Friday review, covering the week that closed the previous Wednesday.", {})],
        [("Net revenue. ", {"bold": True}),
         ("Recognised on invoice creation, dated to service date. Fleet channel and the "
          "job_0 field are excluded. Source: fact_nubrakes_channel_market_kpi_daily.", {})],
        [("Average ticket. ", {"bold": True}),
         ("Net revenue ÷ jobs completed, same period.", {})],
        [("Booked ticket. ", {"bold": True}),
         ("Booked revenue ÷ revenue-producing jobs — the value of the work dispatched "
          "to a technician, before they touch it. Left unshaded on the leaderboard "
          "because ranking it would score the schedule, not the person. Average ticket "
          "÷ booked ticket is exactly ticket completion, so the three columns read "
          "left to right as handed, closed, and the spread between them.", {})],
        [("Ticket completion. ", {"bold": True}),
         ("Invoiced revenue ÷ booked revenue, both from the technician feed. Above 100% "
          "means work found on site beyond what was booked; below means jobs closing "
          "short of their quote.", {})],
        [("Rotor attach. ", {"bold": True}),
         ("Jobs with a rotor attached ÷ revenue-producing jobs — the basis the rotor "
          "attachment dashboard has always used. The new ops feed also carries a "
          "rotor-eligible count, and attach measured against that runs about 7 points "
          "higher because eligibility screens jobs out first. Eligibility itself varies "
          "sharply by market (57% to 97% of jobs), so the eligible basis is not "
          "comparable across markets. Switch with ROTOR_DENOMINATOR in compute.py.", {})],
        [("Technician efficiency. ", {"bold": True}),
         ("Invoiced hours ÷ clocked hours. Clocked hours are not in the published feed "
          "and reach Ops on Friday afternoon, after this deck ships, so the figure shows "
          "a dash until a file is supplied. Drop data/clocked_hours.csv with columns "
          "technician, week_start, clocked_hours and the column fills itself.", {})],
        [("Rev-producing jobs per tech day. ", {"bold": True}),
         ("Completed jobs that carry an invoice, ÷ technician days worked. Jobs without "
          "an invoice are excluded throughout the deck. A partial day counts as a full "
          "day, so this is a productivity measure, not a capacity one. "
          "Source: fact_technician_operations_daily, which supersedes the older "
          "technician feed — its revenue ties to the channel feed to the dollar each "
          "week, where the older one drifts up to 2%.", {})],
        [("Service call rate. ", {"bold": True}),
         ("Service calls ÷ (service calls + revenue-producing jobs). A rising rate means "
          "more dispatches that do not convert to repair revenue.", {})],
        [("Revenue bridge. ", {"bold": True}),
         ("Week-over-week revenue change split into capacity (change in tech days worked, "
          "held at last week's revenue per day) and productivity (change in revenue per "
          "tech day, at this week's days). The two components sum to the total.", {})],
        [("Outliers. ", {"bold": True}),
         ("A market-metric is flagged when it sits more than 1.5 standard deviations from "
          "its trailing 8-week mean AND clears a materiality floor ($4,000 revenue, $25 "
          "ticket, 0.25 jobs per tech day, 3pp service call rate). Both gates must clear, "
          "so a large swing on a small base does not dominate the read-out.", {})],
        [("Month pace. ", {"bold": True}),
         ("Month-to-date revenue ÷ elapsed days × days in month, against last month's "
          "actual. Taken through the freshest day in the feed, which runs ahead of the "
          "reported operations week. This is a straight-line extrapolation, not a "
          "Finance estimate. When an LE is loaded, pace is measured against it instead.",
          {})],
        [("Labor %. ", {"bold": True}),
         ("The P&L Cost of Sales labor line ÷ net revenue, by market by month. Labor "
          "dollars come from the EGP model's financial_model tab (G/L basis); the "
          "denominator is the channel feed, which matches the model's net sales to the "
          "dollar. Salaried ops-admin pay sits in Market G&A and SG&A and is not "
          "separable, so this is store and technician labor.", {})],
        [("Technician leaderboard. ", {"bold": True}),
         ("A trailing four weeks, not one — a single week leaves a technician with too "
          "few jobs to rank fairly. Shaded columns rank technicians within their own "
          "market; volume columns (days, revenue, jobs, invoiced hours) are left "
          "unshaded because ranking them just re-ranks headcount.", {})],
        [("The ▲ and ▼ markers. ", {"bold": True}),
         ("A technician whose trailing four weeks have moved against their OWN prior "
          "eight — average ticket by more than 12%, or service call rate by more than 5 "
          "points, in either direction. The same bar applies both ways, so a recovery "
          "is not held to a stricter standard than a slide. Shading answers who is "
          "weakest in this market; the marker answers what has changed for that person. "
          "A strong technician having a bad month earns a ▼; a consistently weak one "
          "does not, and a weak one who is turning it around earns a ▲. Where someone "
          "has moved both ways at once the marker shows the adverse direction, because "
          "that is what needs a conversation, and the footnote lists both. Every marker "
          "carries its reason under the table.", {})],
        [("Profitability and labor. ", {"bold": True}),
         ("The Profitability & Unit Economics slide is out of the weekly deck for now. "
          "The EGP model and the monthly labor file are still loaded, so it can be "
          "restored without rework.", {})],
        [("Technician history. ", {"bold": True}),
         ("The technician feed begins 16 March 2026, so tech-day, ticket-completion and "
          "rotor-attach metrics carry no year-over-year comparison until March 2027. "
          "Revenue history runs to October 2024.", {})],
    ]
    txbox(slide, MARGIN, 1.16, CONTENT_W, 5.9,
          [[("•  ", {"color": BLUE, "bold": True})] + l for l in lines],
          size=10, color=INK, space_after=7)
    txbox(slide, MARGIN, 7.08, CONTENT_W, 0.24,
          f"Generated {m['generated'].replace('T', ' ')} · feeds current through "
          f"{fmt_date(m['feed_last_day']['ds01'])}",
          size=8, color=MUTED, italic=True)
    return slide


def build(metrics_path, out_dir):
    with open(metrics_path) as f:
        m = json.load(f)
    m["_out"] = out_dir
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(SLIDE_W), Inches(SLIDE_H)

    exec_summary(prs, m)
    # Profitability & unit economics is built but not shipped — the EGP model and the
    # labor file are still loaded, so restoring the slide is flipping this flag.
    if INCLUDE_PROFITABILITY and m.get("egp"):
        profitability(prs, m)
    for region in m["regions"]:
        region_slide(prs, m, region)
    # Market slides sit behind the core four, ordered by region then revenue, so each
    # manager's markets stay together.
    for region, cfg in m["regions"].items():
        for market in sorted(cfg["markets"], key=lambda k: -(m["markets"][k]["revenue"] or 0)):
            market_slide(prs, m, market)
    title_note(prs, m)

    wk = dt.date.fromisoformat(m["ops_week"]["end"])
    path = os.path.join(out_dir, f"NuBrakes_OBR_{wk:%Y-%m-%d}.pptx")
    prs.save(path)
    print(f"wrote {path} ({len(prs.slides.__iter__.__self__._sldIdLst)} slides)")
    return path


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--metrics", default="out/metrics.json")
    ap.add_argument("--out", default="out")
    a = ap.parse_args()
    build(a.metrics, a.out)
