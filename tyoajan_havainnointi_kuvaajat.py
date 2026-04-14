"""
Työajan havainnointi – Kuvaajageneraattori
========================================

Lukee yhden tai useamman Excel-tiedoston "Havainnot"-välilehdeltä
ja piirtää jokaiselle henkilölle omat kuvaajat:

  1. Aikajanakaavio  – Tekemisaika yläpuolella, Apuaika alapuolella,
                       Häiriöaika ja muu aika omilla väreillään.
  2. Työneräkaavio   – Tekemisaika yläpuolella, muut ajat alapuolella.
                       Työnerät esitetään omilla väreillään palkkeina.
                       Lisäksi näytetään työneräkohtaiset jaksojen pituustilastot.
  3. Yhteenvetokuvaaja – Aikalajien prosenttiosuudet päivittäin.

Käyttö:
  Streamlit:   streamlit run tyoajan_havainnointi_kuvaajat.py
  Paikallinen: python tyoajan_havainnointi_kuvaajat.py

Vaatimukset (requirements.txt):
  openpyxl
  matplotlib
  streamlit
  pandas
"""

import math
import sys
from collections import defaultdict
from datetime import datetime

import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import openpyxl
import pandas as pd

# ── Tunnistetaan ajoympäristö ──────────────────────────────────────────────

try:
    import streamlit as st
    STREAMLIT = True
except ImportError:
    STREAMLIT = False

# ── Asetukset ──────────────────────────────────────────────────────────────

GAP_BETWEEN_DAYS = 20

MARKER_SIZE_HAIRIO = 600
MARKER_SIZE_MUU = 500

COLORS = {
    "Tekemisaika": "#1976D2",
    "Apuaika":     "#FB8C00",
    "Häiriöaika":  "#E53935",
    "Muu":         "#8E24AA",
    "Valmiusaika": "#43A047",
    "Taukoaika":   "#757575",
    "Tuntematon":  "#9E9E9E",
}

SUMMARY_CATS = [
    "Tekemisaika", "Apuaika", "Häiriöaika",
    "Valmiusaika", "Taukoaika", "Muu", "Tuntematon",
]

PERSON_COLUMNS = {
    2: "Henkilö 1",  # C
    3: "Henkilö 2",  # D
    4: "Henkilö 3",  # E
    5: "Henkilö 4",  # F
}

# ── Luokittelufunktio ──────────────────────────────────────────────────────

def classify_code(code) -> str:
    """Palauttaa aikalajin havaintokoodin perusteella."""
    try:
        c = float(str(code).replace(",", "."))
    except (ValueError, TypeError):
        return "Tuntematon"

    if 1 <= c < 40:
        return "Tekemisaika"
    if 40 <= c < 50:
        return "Apuaika"
    if c == 50:
        return "Valmiusaika"
    if 51 <= c <= 55:
        return "Häiriöaika"
    if 56 <= c <= 59:
        return "Muu"
    if c >= 60:
        return "Taukoaika"
    return "Tuntematon"


# ── Datan luku ─────────────────────────────────────────────────────────────

def read_batch_catalog(wb) -> dict:
    """
    Lukee välilehden "Eräluettelo" ja palauttaa sanakirjan:
      {code: {"name": "...", "category": "..."}}

    Mukaan otetaan vain ne erät, joilla on erän nimi.
    """
    if "Eräluettelo" not in wb.sheetnames:
        return {}

    ws = wb["Eräluettelo"]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return {}

    header_idx = None
    col_code = None
    col_name = None
    col_cat = None

    for i, row in enumerate(rows):
        if not row:
            continue
        normalized = [str(v).strip() if v is not None else "" for v in row]
        if "Eränro" in normalized and "Erän nimi" in normalized:
            header_idx = i
            col_code = normalized.index("Eränro")
            col_name = normalized.index("Erän nimi")
            col_cat = normalized.index("Aikalaji") if "Aikalaji" in normalized else None
            break

    if header_idx is None:
        return {}

    batch_catalog = {}
    for row in rows[header_idx + 1:]:
        if not row:
            continue

        code = row[col_code] if col_code < len(row) else None
        name = row[col_name] if col_name < len(row) else None
        cat = row[col_cat] if (col_cat is not None and col_cat < len(row)) else None

        if code is None or name is None:
            continue

        name_str = str(name).strip()
        if not name_str:
            continue

        try:
            code_num = float(str(code).replace(",", "."))
            code_key = int(code_num) if code_num.is_integer() else code_num
        except Exception:
            code_key = str(code).strip()

        batch_catalog[code_key] = {
            "name": name_str,
            "category": str(cat).strip() if cat is not None else "",
        }

    return batch_catalog


def normalize_code_key(code):
    try:
        code_num = float(str(code).replace(",", "."))
        return int(code_num) if code_num.is_integer() else code_num
    except Exception:
        return str(code).strip()


def read_file(filepath) -> dict:
    """
    Lukee yhden Excel-tiedoston havaintodatan.

    Palauttaa muodon:
    {
        "date": <date|None>,
        "persons": {
            "Henkilö 1": [obs, obs, ...],
            ...
        },
        "batch_catalog": {code: {"name": ..., "category": ...}}
    }
    """
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
    except Exception as e:
        raise IOError(f"Tiedoston avaaminen epäonnistui: {e}")

    if "Havainnot" not in wb.sheetnames:
        wb.close()
        raise ValueError("Tiedostosta puuttuu välilehti 'Havainnot'.")

    batch_catalog = read_batch_catalog(wb)

    ws = wb["Havainnot"]
    rows = list(ws.iter_rows(values_only=True))

    date_val = rows[2][5] if len(rows) > 2 and len(rows[2]) > 5 else None
    meas_date = date_val.date() if isinstance(date_val, datetime) else None

    if len(rows) < 6:
        wb.close()
        return {"date": meas_date, "persons": {}, "batch_catalog": batch_catalog}

    person_observations = defaultdict(list)

    for row in rows[5:]:
        if len(row) < 2 or row[0] is None or row[1] is None:
            continue

        try:
            h = int(row[0])
            m = int(row[1])
        except (TypeError, ValueError):
            continue

        for col_idx, person_name in PERSON_COLUMNS.items():
            code = row[col_idx] if col_idx < len(row) else None
            if code is None or str(code).strip() == "":
                continue

            person_observations[person_name].append({
                "hour": h,
                "minute": m,
                "code": code,
                "code_key": normalize_code_key(code),
                "category": classify_code(code),
            })

    wb.close()
    return {
        "date": meas_date,
        "persons": dict(person_observations),
        "batch_catalog": batch_catalog,
    }


# ── Apufunktiot ────────────────────────────────────────────────────────────

def build_segments(series: list, category: str) -> list:
    segs = []
    in_seg = False
    seg_start = None
    prev_x = None

    for obs in series:
        if obs["category"] == category:
            if not in_seg:
                seg_start = obs["x"]
                in_seg = True
            prev_x = obs["x"]
        else:
            if in_seg:
                segs.append((seg_start, prev_x + 1))
                in_seg = False

    if in_seg and prev_x is not None:
        segs.append((seg_start, prev_x + 1))

    return segs


def minutes_to_label(abs_minutes: float) -> str:
    h = int(abs_minutes // 60) % 24
    m = int(abs_minutes % 60)
    return f"{h:02d}:{m:02d}"


def build_day_info(datasets: list) -> list:
    day_info = []
    x_offset = 0

    for ds in datasets:
        obs = ds["observations"]
        if not obs:
            continue

        obs = sorted(obs, key=lambda o: (o["hour"], o["minute"]))
        start_abs = obs[0]["hour"] * 60 + obs[0]["minute"]
        series = []

        for o in obs:
            abs_min = o["hour"] * 60 + o["minute"]
            x = x_offset + (abs_min - start_abs)
            series.append({**o, "x": x, "abs_min": abs_min})

        end_x = series[-1]["x"]
        day_info.append({
            "date": ds["date"],
            "x_start": x_offset,
            "x_end": end_x + 1,
            "x_mid": (x_offset + end_x) / 2,
            "start_abs": start_abs,
            "series": series,
        })
        x_offset = end_x + 1 + GAP_BETWEEN_DAYS

    return day_info


def build_person_day_infos(file_datasets: list):
    """
    Henkilöt yhdistetään tiedostojen välillä sarakepaikan perusteella:
      C = Henkilö 1, D = Henkilö 2, E = Henkilö 3, F = Henkilö 4
    """
    person_days = defaultdict(list)
    merged_batch_catalog = {}

    for ds in file_datasets:
        date_val = ds["date"]
        for code_key, info in ds.get("batch_catalog", {}).items():
            if code_key not in merged_batch_catalog:
                merged_batch_catalog[code_key] = info

        for person_name, observations in ds["persons"].items():
            if not observations:
                continue
            person_days[person_name].append({
                "date": date_val,
                "observations": observations,
            })

    for person_name in person_days:
        person_days[person_name].sort(key=lambda d: d["date"] or datetime.min.date())

    person_day_infos = {
        person_name: build_day_info(day_datasets)
        for person_name, day_datasets in person_days.items()
    }
    return person_day_infos, merged_batch_catalog


def get_xticks_for_day_info(day_info: list):
    tick_positions = []
    tick_labels = []
    end_tick_positions = set()

    for day in day_info:
        duration = day["x_end"] - day["x_start"]
        for offset in range(0, int(duration) + 1, 60):
            tick_positions.append(day["x_start"] + offset)
            tick_labels.append(minutes_to_label(day["start_abs"] + offset))

        last_series = day["series"][-1]
        end_abs = last_series["abs_min"]
        end_x = last_series["x"]

        if end_x not in tick_positions:
            tick_positions.append(end_x)
            tick_labels.append(minutes_to_label(end_abs))
            end_tick_positions.add(end_x)

    ticks_sorted = sorted(zip(tick_positions, tick_labels), key=lambda t: t[0])
    if not ticks_sorted:
        return [], [], set()

    tick_positions = [t[0] for t in ticks_sorted]
    tick_labels = [t[1] for t in ticks_sorted]
    return tick_positions, tick_labels, end_tick_positions


def sort_code_key(code):
    try:
        return float(str(code).replace(",", "."))
    except Exception:
        return str(code)


def format_batch_label(code_key, batch_catalog):
    info = batch_catalog.get(code_key, {})
    name = str(info.get("name", "")).strip()
    code_text = str(code_key).replace(".0", "")
    return f"{code_text} – {name}" if name else f"Työnerä {code_text}"


def get_code_runs(series):
    if not series:
        return []

    runs = []
    current = series[0]
    run_start_x = current["x"]
    prev_x = current["x"]
    run_len = 1

    for obs in series[1:]:
        if obs["code_key"] == current["code_key"]:
            prev_x = obs["x"]
            run_len += 1
        else:
            runs.append({
                "x0": run_start_x,
                "x1": prev_x + 1,
                "length": run_len,
                "category": current["category"],
                "code": current["code"],
                "code_key": current["code_key"],
            })
            current = obs
            run_start_x = obs["x"]
            prev_x = obs["x"]
            run_len = 1

    runs.append({
        "x0": run_start_x,
        "x1": prev_x + 1,
        "length": run_len,
        "category": current["category"],
        "code": current["code"],
        "code_key": current["code_key"],
    })
    return runs


def compute_batch_run_statistics(day_info: list, batch_catalog: dict) -> pd.DataFrame:
    """
    Laskee työneräkohtaiset yhtäjaksoisten jaksojen pituustilastot.

    Mukaan otetaan vain:
    - sellaiset erät, joilla on nimi Eräluettelo-välilehdellä
    - tekemisaikaan kuuluvat jaksot
    - taulukko järjestetään keskiarvon mukaan laskevaan järjestykseen
    """
    lengths_by_code = defaultdict(list)

    for day in day_info:
        for run in get_code_runs(day["series"]):
            code_key = run["code_key"]
            if code_key not in batch_catalog:
                continue
            if run["category"] != "Tekemisaika":
                continue
            lengths_by_code[code_key].append(run["length"])

    rows = []
    for code_key in lengths_by_code.keys():
        lengths = lengths_by_code[code_key]
        if not lengths:
            continue

        info = batch_catalog.get(code_key, {})
        mean_val = sum(lengths) / len(lengths)
        rows.append({
            "Eränro": str(code_key).replace(".0", ""),
            "Erän nimi": info.get("name", ""),
            "Jaksoja (n)": len(lengths),
            "Min (min)": int(min(lengths)),
            "Max (min)": int(max(lengths)),
            "Keskiarvo (min)": round(mean_val, 2),
            "Keskihajonta": round(pd.Series(lengths).std(ddof=1), 2) if len(lengths) > 1 else 0.0,
        })

    if not rows:
        return pd.DataFrame(columns=[
            "Eränro", "Erän nimi", "Jaksoja (n)", "Min (min)",
            "Max (min)", "Keskiarvo (min)", "Keskihajonta"
        ])

    df = pd.DataFrame(rows)
    df = df.sort_values(by=["Keskiarvo (min)", "Max (min)", "Eränro"], ascending=[False, False, True])
    df = df.reset_index(drop=True)
    return df


# ── Kuvaajien piirto ───────────────────────────────────────────────────────

def make_chart1(day_info: list, person_name: str = ""):
    fig, ax = plt.subplots(figsize=(20, 4))
    fig.patch.set_facecolor("#F5F5F5")
    ax.set_facecolor("#FAFAFA")

    first_day = True
    for day in day_info:
        series = day["series"]

        for x0, x1 in build_segments(series, "Tekemisaika"):
            ax.broken_barh([(x0, x1 - x0)], (0, 1), facecolor=COLORS["Tekemisaika"], alpha=0.85,
                           label="Tekemisaika" if first_day else "")

        for x0, x1 in build_segments(series, "Apuaika"):
            ax.broken_barh([(x0, x1 - x0)], (-1, 1), facecolor=COLORS["Apuaika"], alpha=0.85,
                           label="Apuaika" if first_day else "")

        for x0, x1 in build_segments(series, "Valmiusaika"):
            ax.broken_barh([(x0, x1 - x0)], (0, 1), facecolor=COLORS["Valmiusaika"], alpha=0.85,
                           label="Valmiusaika" if first_day else "")

        for x0, x1 in build_segments(series, "Taukoaika"):
            ax.broken_barh([(x0, x1 - x0)], (-1, 1), facecolor=COLORS["Taukoaika"], alpha=0.70,
                           label="Taukoaika" if first_day else "")

        hairio = [o["x"] + 0.5 for o in series if o["category"] == "Häiriöaika"]
        if hairio:
            ax.scatter(hairio, [0] * len(hairio), marker=(10, 1, 0), s=MARKER_SIZE_HAIRIO,
                       color=COLORS["Häiriöaika"], zorder=6, label="Häiriöaika" if first_day else "",
                       linewidths=0.5, edgecolors="white")

        muu = [o["x"] + 0.5 for o in series if o["category"] == "Muu"]
        if muu:
            ax.scatter(muu, [0] * len(muu), marker="o", s=MARKER_SIZE_MUU,
                       color=COLORS["Muu"], zorder=6, label="Muu" if first_day else "",
                       linewidths=0.8, edgecolors="white")

        first_day = False

    tick_positions, tick_labels, end_tick_positions = get_xticks_for_day_info(day_info)
    ax.set_xticks(tick_positions)
    ax.set_xticklabels(tick_labels, rotation=45, ha="right", fontsize=8)

    for label, pos in zip(ax.get_xticklabels(), tick_positions):
        if pos in end_tick_positions:
            label.set_color("#C62828")
            label.set_fontweight("bold")
        else:
            label.set_color("#333333")
            label.set_fontweight("normal")

    for day in day_info:
        end_x = day["series"][-1]["x"]
        ax.axvline(end_x, color="#C62828", linestyle=":", linewidth=1.2, zorder=4)

    for day in day_info:
        date_label = day["date"].strftime("%d.%m.%Y") if day["date"] else "?"
        x_start = day["x_start"]
        ax.axvline(x_start, color="#888888", linestyle="--", linewidth=1.2, zorder=4)
        ax.text(x_start + 1.0, -1.45, date_label, rotation=90, rotation_mode="anchor",
                ha="left", va="bottom", fontsize=9, fontweight="bold", color="#444444", zorder=5)

    if len(day_info) > 1:
        for i in range(len(day_info) - 1):
            gap_x = (day_info[i]["x_end"] + day_info[i + 1]["x_start"]) / 2
            ax.axvline(gap_x, color="#BBBBBB", linestyle="--", linewidth=1.0, zorder=2)

    ax.axhline(0, color="#333333", linewidth=1.2, zorder=3)
    ax.set_ylim(-1.6, 2.0)
    ax.set_yticks([-0.5, 0.5])
    ax.set_yticklabels(["Apuaika", "Tekemisaika"], fontsize=9)
    ax.tick_params(axis="y", length=0)

    for spine in ["top", "right"]:
        ax.spines[spine].set_visible(False)
    ax.spines["left"].set_color("#CCCCCC")
    ax.spines["bottom"].set_color("#CCCCCC")
    ax.grid(axis="x", color="#E0E0E0", linestyle="-", linewidth=0.5, zorder=0)

    legend_handles = [
        mpatches.Patch(color=COLORS["Tekemisaika"], label="Tekemisaika"),
        mpatches.Patch(color=COLORS["Apuaika"], label="Apuaika"),
        mpatches.Patch(color=COLORS["Valmiusaika"], label="Valmiusaika"),
        mpatches.Patch(color=COLORS["Taukoaika"], label="Taukoaika"),
        plt.Line2D([0], [0], marker=(10, 1, 0), color="w", markerfacecolor=COLORS["Häiriöaika"],
                   markersize=14, label="Häiriöaika"),
        plt.Line2D([0], [0], marker="o", color="w", markerfacecolor=COLORS["Muu"],
                   markersize=12, label="Muu"),
        plt.Line2D([0], [0], color="#888888", linestyle="--", linewidth=1.5, label="Mittauksen aloitus"),
        plt.Line2D([0], [0], color="#C62828", linestyle=":", linewidth=1.5, label="Mittauksen päättyminen"),
    ]
    ax.legend(handles=legend_handles, loc="lower right", framealpha=0.9, fontsize=8, ncol=4)

    title = "Aikalajit ajan suhteen – minuuttihavainnot"
    if person_name:
        title += f" ({person_name})"
    ax.set_title(title, fontsize=13, fontweight="bold", pad=30)
    ax.set_xlabel("Kellonaika", fontsize=9)
    plt.tight_layout()
    return fig


def make_chart2(day_info: list, person_name: str = ""):
    n_days = len(day_info)
    fig, axes = plt.subplots(1, n_days, figsize=(max(5 * n_days, 6), 5), sharey=False, squeeze=False)
    fig.patch.set_facecolor("#F5F5F5")

    title = "Aikalajien osuudet päivittäin"
    if person_name:
        title += f" ({person_name})"
    fig.suptitle(title, fontsize=13, fontweight="bold", y=1.02)

    for col_i, day in enumerate(day_info):
        ax = axes[0][col_i]
        ax.set_facecolor("#FAFAFA")
        total = len(day["series"])

        pcts = [
            100 * sum(1 for o in day["series"] if o["category"] == cat) / total if total > 0 else 0
            for cat in SUMMARY_CATS
        ]

        bars = ax.barh(
            SUMMARY_CATS,
            pcts,
            color=[COLORS.get(c, "#BDBDBD") for c in SUMMARY_CATS],
            edgecolor="white",
            height=0.65,
            alpha=0.9,
        )

        for bar, pct in zip(bars, pcts):
            if pct > 0.5:
                ax.text(bar.get_width() + 0.5, bar.get_y() + bar.get_height() / 2,
                        f"{pct:.1f} %", va="center", ha="left", fontsize=9, color="#333333")

        label = day["date"].strftime("%d.%m.%Y") if day["date"] else "?"
        ax.set_title(label, fontsize=11, fontweight="bold")
        ax.set_xlabel("Osuus (%)", fontsize=9)
        ax.set_xlim(0, max(max(pcts) * 1.25 + 2, 10))

        for spine in ["top", "right"]:
            ax.spines[spine].set_visible(False)
        ax.spines["left"].set_color("#CCCCCC")
        ax.spines["bottom"].set_color("#CCCCCC")
        ax.grid(axis="x", color="#E0E0E0", linestyle="-", linewidth=0.5)

        ax.text(0.02, 0.98, f"n = {total} havaintoa / ≈ {total} min",
                transform=ax.transAxes, ha="left", va="top", fontsize=8, color="#777777")

    plt.tight_layout()
    return fig


def make_chart3(day_info: list, batch_catalog: dict, person_name: str = ""):
    """
    Työneräkaavio:
    - Tekemisaika ja valmiusaika yläpuolella
    - Apuaika alapuolella työnerän värillä
    - Häiriöaika, muu, taukoaika ja tuntematon alapuolella kiinteillä väreillä
    - Legendassa näkyvät työnerän numero + nimi, jos nimi löytyy Eräluettelo-välilehdeltä
    """
    def build_code_colors(day_info_local):
        unique_codes = []
        seen = set()

        for day in day_info_local:
            for obs in day["series"]:
                if obs["code_key"] not in batch_catalog:
                    continue
                if obs["category"] not in ("Tekemisaika", "Apuaika", "Valmiusaika"):
                    continue
                code_key = obs["code_key"]
                if code_key not in seen:
                    seen.add(code_key)
                    unique_codes.append(code_key)

        unique_codes = sorted(unique_codes, key=sort_code_key)
        cmap = plt.get_cmap("tab20")
        color_map = {}
        for i, code_key in enumerate(unique_codes):
            color_map[code_key] = cmap(i % 20)
        return color_map, unique_codes

    code_colors, unique_codes = build_code_colors(day_info)

    fig, ax = plt.subplots(figsize=(20, 4.2))
    fig.patch.set_facecolor("#F5F5F5")
    ax.set_facecolor("#FAFAFA")

    used_codes = set()

    for day in day_info:
        runs = get_code_runs(day["series"])

        for run in runs:
            x0 = run["x0"]
            width = run["x1"] - run["x0"]
            cat = run["category"]
            code_key = run["code_key"]

            if cat in ("Tekemisaika", "Valmiusaika") and code_key in batch_catalog:
                color = code_colors.get(code_key, COLORS["Tekemisaika"])
                used_codes.add(code_key)
                ax.broken_barh([(x0, width)], (0, 1), facecolor=color, edgecolor="white",
                               linewidth=0.6, alpha=0.95, zorder=3)
            elif cat == "Apuaika" and code_key in batch_catalog:
                color = code_colors.get(code_key, COLORS["Apuaika"])
                used_codes.add(code_key)
                ax.broken_barh([(x0, width)], (-1, 1), facecolor=color, edgecolor="white",
                               linewidth=0.6, alpha=0.95, zorder=3)
            elif cat in ("Häiriöaika", "Muu", "Taukoaika", "Tuntematon") or code_key not in batch_catalog:
                ax.broken_barh([(x0, width)], (-1, 1), facecolor=COLORS.get(cat, COLORS["Tuntematon"]),
                               edgecolor="white", linewidth=0.6, alpha=0.95, zorder=3)

    tick_positions, tick_labels, end_tick_positions = get_xticks_for_day_info(day_info)
    ax.set_xticks(tick_positions)
    ax.set_xticklabels(tick_labels, rotation=45, ha="right", fontsize=8)

    for label, pos in zip(ax.get_xticklabels(), tick_positions):
        if pos in end_tick_positions:
            label.set_color("#C62828")
            label.set_fontweight("bold")
        else:
            label.set_color("#333333")
            label.set_fontweight("normal")

    for day in day_info:
        end_x = day["series"][-1]["x"]
        ax.axvline(end_x, color="#C62828", linestyle=":", linewidth=1.2, zorder=4)

    for day in day_info:
        date_label = day["date"].strftime("%d.%m.%Y") if day["date"] else "?"
        x_start = day["x_start"]
        ax.axvline(x_start, color="#888888", linestyle="--", linewidth=1.2, zorder=4)
        ax.text(x_start + 1.0, -1.45, date_label, rotation=90, rotation_mode="anchor",
                ha="left", va="bottom", fontsize=9, fontweight="bold", color="#444444", zorder=5)

    if len(day_info) > 1:
        for i in range(len(day_info) - 1):
            gap_x = (day_info[i]["x_end"] + day_info[i + 1]["x_start"]) / 2
            ax.axvline(gap_x, color="#BBBBBB", linestyle="--", linewidth=1.0, zorder=2)

    ax.axhline(0, color="#333333", linewidth=1.2, zorder=3)
    ax.set_ylim(-1.6, 2.0)
    ax.set_yticks([-0.5, 0.5])
    ax.set_yticklabels(["Apuaika, häiriöaika ja muu aika", "Tekemisaika"], fontsize=9)
    ax.tick_params(axis="y", length=0)
    ax.set_xlabel("Kellonaika", fontsize=9)

    for spine in ["top", "right"]:
        ax.spines[spine].set_visible(False)
    ax.spines["left"].set_color("#CCCCCC")
    ax.spines["bottom"].set_color("#CCCCCC")
    ax.grid(axis="x", color="#E0E0E0", linestyle="-", linewidth=0.5, zorder=0)

    code_handles = [
        mpatches.Patch(color=code_colors[code_key], label=format_batch_label(code_key, batch_catalog))
        for code_key in unique_codes if code_key in used_codes
    ]
    fixed_handles = [
        mpatches.Patch(color=COLORS["Häiriöaika"], label="Häiriöaika"),
        mpatches.Patch(color=COLORS["Muu"], label="Muu"),
        mpatches.Patch(color=COLORS["Taukoaika"], label="Taukoaika"),
        mpatches.Patch(color=COLORS["Tuntematon"], label="Tuntematon / nimeämätön erä"),
        plt.Line2D([0], [0], color="#888888", linestyle="--", linewidth=1.5, label="Mittauksen aloitus"),
        plt.Line2D([0], [0], color="#C62828", linestyle=":", linewidth=1.5, label="Mittauksen päättyminen"),
    ]
    legend_handles = code_handles + fixed_handles

    ax.legend(handles=legend_handles, loc="upper center", bbox_to_anchor=(0.5, -0.22),
              framealpha=0.95, fontsize=8,
              ncol=min(4, max(2, (len(legend_handles) + 2) // 3)))

    title = "Työneräkaavio – työnerät väreillä, muut ajat alapuolella"
    if person_name:
        title += f" ({person_name})"
    ax.set_title(title, fontsize=12, fontweight="bold", pad=20)

    plt.tight_layout()
    return fig


# ── Käyttöliittymät ────────────────────────────────────────────────────────

def render_person_sections(person_day_infos: dict, batch_catalog: dict, ui_mode: str = "streamlit"):
    if not person_day_infos:
        if ui_mode == "streamlit":
            st.warning("Valituissa tiedostoissa ei ole havaintodataa.")
            st.stop()
        else:
            print("Valituissa tiedostoissa ei ole havaintodataa.")
            sys.exit(0)

    for person_name in sorted(person_day_infos.keys()):
        day_info = person_day_infos[person_name]
        if not day_info:
            continue

        stats_df = compute_batch_run_statistics(day_info, batch_catalog)

        if ui_mode == "streamlit":
            st.header(f"Henkilö: {person_name}")
            st.subheader("Kuvaaja 1 – Aikajanakaavio")
            st.pyplot(make_chart1(day_info, person_name))

            st.subheader("Kuvaaja 1b – Työneräkaavio")
            st.pyplot(make_chart3(day_info, batch_catalog, person_name))

            st.markdown("**Työneräkohtaiset yhtäjaksoisen tekemisen tilastot**")
            if stats_df.empty:
                st.info("Eräluettelo-välilehdeltä ei löytynyt nimettyjä työneriä, joita olisi havaittu tässä datassa.")
            else:
                st.dataframe(stats_df, use_container_width=True, hide_index=True)

            st.subheader("Kuvaaja 2 – Yhteenveto päivittäin")
            st.pyplot(make_chart2(day_info, person_name))
        else:
            make_chart1(day_info, person_name)
            make_chart3(day_info, batch_catalog, person_name)
            make_chart2(day_info, person_name)
            if not stats_df.empty:
                print(f"\n{person_name} – Työneräkohtaiset yhtäjaksoisen tekemisen tilastot")
                print(stats_df.to_string(index=False))


def run_streamlit():
    st.set_page_config(page_title="Työajan havainnointi", layout="wide")
    st.title("Työajan havainnointi – Kuvaajageneraattori")
    st.write("Lataa yksi tai useampi havainnointi-Excel-tiedosto (.xlsx tai .xlsm).")
    st.write("Jos tiedostossa on useita henkilöitä, jokaiselle henkilösarakkeelle muodostetaan omat kuvaajat.")
    st.write("Tulkinta on kiinteä: C = Henkilö 1, D = Henkilö 2, E = Henkilö 3 ja F = Henkilö 4 kaikissa tiedostoissa.")

    uploaded = st.file_uploader(
        "Valitse Excel-tiedosto(t)",
        type=["xlsx", "xlsm"],
        accept_multiple_files=True,
    )

    if not uploaded:
        st.info("Valitse vähintään yksi tiedosto yllä.")
        st.stop()

    file_datasets = []
    for f in uploaded:
        try:
            file_datasets.append(read_file(f))
        except Exception as e:
            st.error(f"Virhe tiedostossa {f.name}: {e}")

    if not file_datasets:
        st.stop()

    person_day_infos, batch_catalog = build_person_day_infos(file_datasets)
    render_person_sections(person_day_infos, batch_catalog, ui_mode="streamlit")


def run_local():
    import tkinter as tk
    from tkinter import filedialog, messagebox

    root = tk.Tk()
    root.withdraw()
    paths = filedialog.askopenfilenames(
        title="Valitse havainnointi-Excel-tiedosto(t)",
        filetypes=[
            ("Excel-tiedostot", "*.xlsx *.xlsm"),
            ("Kaikki tiedostot", "*.*"),
        ],
    )
    root.destroy()

    if not paths:
        print("Ei valittuja tiedostoja – ohjelma lopetetaan.")
        sys.exit(0)

    file_datasets = []
    errors = []
    for p in paths:
        try:
            file_datasets.append(read_file(p))
        except Exception as e:
            errors.append(f"{p}: {e}")

    if errors:
        root2 = tk.Tk()
        root2.withdraw()
        messagebox.showerror("Virhe tiedostoja luettaessa", "\n".join(errors))
        root2.destroy()

    if not file_datasets:
        sys.exit(1)

    person_day_infos, batch_catalog = build_person_day_infos(file_datasets)
    render_person_sections(person_day_infos, batch_catalog, ui_mode="local")
    plt.show()


if __name__ == "__main__":
    if STREAMLIT:
        run_streamlit()
    else:
        run_local()
