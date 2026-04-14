"""
Työajan havainnointi – Kuvaajageneraattori
==========================================
Lukee yhden tai useamman Excel-tiedoston "Havainnot"-välilehdeltä
ja piirtää kaksi kuvaajaa:

  1. Aikajanakaavio  – Tekemisaika yläpuolella, Apuaika alapuolella,
                       Häiriöaika räjähdyssymbolilla, Muu ympyräsymbolilla.
  2. Yhteenvetokuvaaja – Aikalajien prosenttiosuudet päivittäin.

Käyttö:
  Streamlit:   streamlit run tyoajan_havainnointi_kuvaajat.py
  Paikallinen: python tyoajan_havainnointi_kuvaajat.py  (avaa tkinter-ikkunan)

Vaatimukset (requirements.txt):
  openpyxl
  matplotlib
  streamlit
"""

import sys
from datetime import datetime

import openpyxl
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np

# ── Tunnistetaan ajoympäristö ──────────────────────────────────────────────

try:
    import streamlit as st
    STREAMLIT = True
except ImportError:
    STREAMLIT = False

# ── Asetukset ──────────────────────────────────────────────────────────────

# Pieni väli päivien välissä (minuuteissa)
GAP_BETWEEN_DAYS = 20

# Symbolien koot (scatter-pisteille pinta-ala pisteinä²)
MARKER_SIZE_HAIRIO = 600   # räjähdyssymboli
MARKER_SIZE_MUU    = 500   # ympyräsymboli

# Värit aikalajille
COLORS = {
    "Tekemisaika": "#1976D2",   # sininen
    "Apuaika":     "#FB8C00",   # oranssi
    "Häiriöaika":  "#E53935",   # punainen
    "Muu":         "#8E24AA",   # violetti
    "Valmiusaika": "#43A047",   # vihreä
    "Taukoaika":   "#757575",   # harmaa
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

def read_file(filepath) -> dict:
    """
    Lukee yhden Excel-tiedoston havaintodatan.
    filepath voi olla tiedostopolku (str) tai Streamlitin UploadedFile-objekti.
    """
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
    except Exception as e:
        raise IOError(f"Tiedoston avaaminen epäonnistui: {e}")

    ws = wb["Havainnot"]
    rows = list(ws.iter_rows(values_only=True))

    # Päivämäärä rivillä 3 (indeksi 2), sarakkeessa F (indeksi 5)
    date_val = rows[2][5] if len(rows) > 2 else None
    meas_date = date_val.date() if isinstance(date_val, datetime) else None

    # Havaintodata alkaa riviltä 6 (indeksi 5)
    # Sarakkeet: A=tunti, B=minuutti, C=Henkilö1, D=Henkilö2, ...
    observations = []
    for row in rows[5:]:
        if row[0] is None or row[1] is None:
            continue
        try:
            h = int(row[0])
            m = int(row[1])
        except (TypeError, ValueError):
            continue

        code = row[2]
        if code is not None:
            observations.append({
                "hour":     h,
                "minute":   m,
                "code":     code,
                "category": classify_code(code),
            })

    wb.close()
    return {"date": meas_date, "observations": observations}

# ── Apufunktiot ────────────────────────────────────────────────────────────

def build_segments(series: list, category: str) -> list:
    """
    Ryhmittelee peräkkäiset saman aikalajin havainnot yhtenäisiksi jaksoiksi.
    Palauttaa listan (x_alku, x_loppu) -pareista.
    """
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
    """Laskee x-koordinaatit kullekin päivälle ja rakentaa day_info-listan."""
    day_info = []
    x_offset = 0

    for ds in datasets:
        obs = ds["observations"]
        if not obs:
            continue

        start_abs = obs[0]["hour"] * 60 + obs[0]["minute"]
        series = []
        for o in obs:
            abs_min = o["hour"] * 60 + o["minute"]
            x = x_offset + (abs_min - start_abs)
            series.append({**o, "x": x, "abs_min": abs_min})

        end_x = series[-1]["x"]
        day_info.append({
            "date":      ds["date"],
            "x_start":   x_offset,
            "x_end":     end_x + 1,
            "x_mid":     (x_offset + end_x) / 2,
            "start_abs": start_abs,
            "series":    series,
        })
        x_offset = end_x + 1 + GAP_BETWEEN_DAYS

    return day_info

# ── Kuvaajien piirto ───────────────────────────────────────────────────────

def make_chart1(day_info: list):
    """Aikajanakaavio."""
    fig, ax = plt.subplots(figsize=(20, 4))
    fig.patch.set_facecolor("#F5F5F5")
    ax.set_facecolor("#FAFAFA")

    first_day = True
    for day in day_info:
        series = day["series"]

        # Tekemisaika – sininen palkki akselin yläpuolella (y 0 → 1)
        for x0, x1 in build_segments(series, "Tekemisaika"):
            ax.broken_barh(
                [(x0, x1 - x0)], (0, 1),
                facecolor=COLORS["Tekemisaika"], alpha=0.85,
                label="Tekemisaika" if first_day else "",
            )

        # Apuaika – oranssi palkki akselin alapuolella (y -1 → 0)
        for x0, x1 in build_segments(series, "Apuaika"):
            ax.broken_barh(
                [(x0, x1 - x0)], (-1, 1),
                facecolor=COLORS["Apuaika"], alpha=0.85,
                label="Apuaika" if first_day else "",
            )

        # Valmiusaika – vihreä palkki yläpuolella
        for x0, x1 in build_segments(series, "Valmiusaika"):
            ax.broken_barh(
                [(x0, x1 - x0)], (0, 1),
                facecolor=COLORS["Valmiusaika"], alpha=0.85,
                label="Valmiusaika" if first_day else "",
            )

        # Taukoaika – harmaa palkki alapuolella
        for x0, x1 in build_segments(series, "Taukoaika"):
            ax.broken_barh(
                [(x0, x1 - x0)], (-1, 1),
                facecolor=COLORS["Taukoaika"], alpha=0.70,
                label="Taukoaika" if first_day else "",
            )

        # Häiriöaika – räjähdyssymboli (10-kärkinen tähti), y=0 keskeisesti
        hairio = [o["x"] + 0.5 for o in series if o["category"] == "Häiriöaika"]
        if hairio:
            ax.scatter(
                hairio, [0] * len(hairio),
                marker=(10, 1, 0),
                s=MARKER_SIZE_HAIRIO,
                color=COLORS["Häiriöaika"],
                zorder=6,
                label="Häiriöaika" if first_day else "",
                linewidths=0.5,
                edgecolors="white",
            )

        # Muu – ympyräsymboli, y=0 keskeisesti
        muu = [o["x"] + 0.5 for o in series if o["category"] == "Muu"]
        if muu:
            ax.scatter(
                muu, [0] * len(muu),
                marker="o",
                s=MARKER_SIZE_MUU,
                color=COLORS["Muu"],
                zorder=6,
                label="Muu" if first_day else "",
                linewidths=0.8,
                edgecolors="white",
            )

        first_day = False

    # X-akselin kellonajat (tunnin välein)
    tick_positions, tick_labels = [], []
    for day in day_info:
        duration = day["x_end"] - day["x_start"]
        for offset in range(0, int(duration) + 1, 60):
            tick_positions.append(day["x_start"] + offset)
            tick_labels.append(minutes_to_label(day["start_abs"] + offset))

    ax.set_xticks(tick_positions)
    ax.set_xticklabels(tick_labels, rotation=45, ha="right", fontsize=8)

    # Katkoviiva + päivämääräteksti jokaisen päivän aloituskohdassa
    for day in day_info:
        date_label = day["date"].strftime("%-d.%-m.%Y") if day["date"] else "?"
        x_start = day["x_start"]

        # Katkoviiva päivän alkuun
        ax.axvline(x_start, color="#888888", linestyle="--", linewidth=1.2, zorder=4)

        # Päivämääräteksti 90° käännetty, katkoviivan vieressä
        ax.text(
            x_start + 1.0,   # hieman viivan oikealle puolelle
            -1.45,            # lähellä pohjaa, teksti nousee ylöspäin
            date_label,
            rotation=90,
            rotation_mode="anchor",
            ha="left",
            va="bottom",
            fontsize=9,
            fontweight="bold",
            color="#444444",
            zorder=5,
        )

    # Katkoviiva + päivämääräteksti päivien välissä (päivän vaihdon kohta)
    if len(day_info) > 1:
        for i in range(len(day_info) - 1):
            gap_x = (day_info[i]["x_end"] + day_info[i + 1]["x_start"]) / 2
            next_date = day_info[i + 1]["date"]
            date_label = next_date.strftime("%-d.%-m.%Y") if next_date else "?"

            # Katkoviiva päivien väliin
            ax.axvline(gap_x, color="#888888", linestyle="--", linewidth=1.2, zorder=4)

            # Päivämääräteksti 90° käännetty viivan vieressä
            ax.text(
                gap_x + 1.0,
                -1.45,
                date_label,
                rotation=90,
                rotation_mode="anchor",
                ha="left",
                va="bottom",
                fontsize=9,
                fontweight="bold",
                color="#444444",
                zorder=5,
            )

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
        mpatches.Patch(color=COLORS["Apuaika"],     label="Apuaika"),
        mpatches.Patch(color=COLORS["Valmiusaika"], label="Valmiusaika"),
        mpatches.Patch(color=COLORS["Taukoaika"],   label="Taukoaika"),
        plt.Line2D([0], [0], marker=(10, 1, 0), color="w",
                   markerfacecolor=COLORS["Häiriöaika"], markersize=14, label="Häiriöaika"),
        plt.Line2D([0], [0], marker="o", color="w",
                   markerfacecolor=COLORS["Muu"], markersize=12, label="Muu"),
    ]
    ax.legend(handles=legend_handles, loc="lower right",
              framealpha=0.9, fontsize=8, ncol=3)
    ax.set_title("Aikalajit ajan suhteen – minuuttihavainnot",
                 fontsize=13, fontweight="bold", pad=30)
    ax.set_xlabel("Kellonaika", fontsize=9)
    plt.tight_layout()
    return fig


def make_chart2(day_info: list):
    """Yhteenvetokuvaaja – aikalajien prosenttiosuudet päivittäin."""
    SUMMARY_CATS = [
        "Tekemisaika", "Apuaika", "Häiriöaika",
        "Valmiusaika", "Taukoaika", "Muu",
    ]

    n_days = len(day_info)
    fig, axes = plt.subplots(1, n_days, figsize=(5 * n_days, 5),
                             sharey=False, squeeze=False)
    fig.patch.set_facecolor("#F5F5F5")
    fig.suptitle("Aikalajien osuudet päivittäin",
                 fontsize=13, fontweight="bold", y=1.02)

    for col_i, day in enumerate(day_info):
        ax = axes[0][col_i]
        ax.set_facecolor("#FAFAFA")
        total = len(day["series"])

        pcts = [
            100 * sum(1 for o in day["series"] if o["category"] == cat) / total
            if total > 0 else 0
            for cat in SUMMARY_CATS
        ]

        bars = ax.barh(
            SUMMARY_CATS, pcts,
            color=[COLORS.get(c, "#BDBDBD") for c in SUMMARY_CATS],
            edgecolor="white", height=0.65, alpha=0.9,
        )

        for bar, pct in zip(bars, pcts):
            if pct > 0.5:
                ax.text(bar.get_width() + 0.5,
                        bar.get_y() + bar.get_height() / 2,
                        f"{pct:.1f} %",
                        va="center", ha="left", fontsize=9, color="#333333")

        label = day["date"].strftime("%-d.%-m.%Y") if day["date"] else "?"
        ax.set_title(label, fontsize=11, fontweight="bold")
        ax.set_xlabel("Osuus (%)", fontsize=9)
        ax.set_xlim(0, max(pcts) * 1.25 + 2)

        for spine in ["top", "right"]:
            ax.spines[spine].set_visible(False)
        ax.spines["left"].set_color("#CCCCCC")
        ax.spines["bottom"].set_color("#CCCCCC")
        ax.grid(axis="x", color="#E0E0E0", linestyle="-", linewidth=0.5)

        ax.text(0.98, 0.02,
                f"n = {total} havaintoa\n≈ {total} min",
                transform=ax.transAxes, ha="right", va="bottom",
                fontsize=8, color="#777777")

    plt.tight_layout()
    return fig

# ── Streamlit-käyttöliittymä ───────────────────────────────────────────────

def run_streamlit():
    st.set_page_config(page_title="Työajan havainnointi", layout="wide")
    st.title("Työajan havainnointi – Kuvaajageneraattori")
    st.write("Lataa yksi tai useampi havainnointi-Excel-tiedosto (.xlsx).")

    uploaded = st.file_uploader(
        "Valitse Excel-tiedosto(t)",
        type=["xlsx", "xlsm"],
        accept_multiple_files=True,
    )

    if not uploaded:
        st.info("Valitse vähintään yksi tiedosto yllä.")
        st.stop()

    datasets = []
    for f in uploaded:
        try:
            datasets.append(read_file(f))
        except Exception as e:
            st.error(f"Virhe tiedostossa {f.name}: {e}")

    if not datasets:
        st.stop()

    datasets.sort(key=lambda d: d["date"] or datetime.min.date())
    day_info = build_day_info(datasets)

    if not day_info:
        st.warning("Valituissa tiedostoissa ei ole havaintodataa.")
        st.stop()

    st.subheader("Kuvaaja 1 – Aikajanakaavio")
    st.pyplot(make_chart1(day_info))

    st.subheader("Kuvaaja 2 – Yhteenveto päivittäin")
    st.pyplot(make_chart2(day_info))

# ── Paikallinen käyttö (tkinter) ───────────────────────────────────────────

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

    datasets = []
    errors = []
    for p in paths:
        try:
            datasets.append(read_file(p))
        except Exception as e:
            errors.append(str(e))

    if errors:
        root2 = tk.Tk()
        root2.withdraw()
        messagebox.showerror("Virhe tiedostoja luettaessa", "\n".join(errors))
        root2.destroy()
    if not datasets:
        sys.exit(1)

    datasets.sort(key=lambda d: d["date"] or datetime.min.date())
    day_info = build_day_info(datasets)

    if not day_info:
        print("Ei havaintodataa – ohjelma lopetetaan.")
        sys.exit(0)

    make_chart1(day_info)
    make_chart2(day_info)
    plt.show()

# ── Käynnistys ─────────────────────────────────────────────────────────────

if STREAMLIT:
    run_streamlit()
else:
    run_local()
