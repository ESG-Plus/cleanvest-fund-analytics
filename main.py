# -*- coding: utf-8 -*-
"""
CLEANVEST Fund Stats — Streamlit replica of the Excel 'Stats' sheet
===================================================================
Reproduces every block of the Stats sheet from DataUnique + DataAllTranches,
plus new analyses (top/worst funds, averages by type, pie charts) and a
built-in verification tab that flags any mismatch against the Stats sheet.

Run:  poetry run streamlit run main.py
Deps: streamlit pandas numpy openpyxl plotly   (pandas/numpy come with streamlit)

Positiv/Negativ rule (reverse-engineered from the Stats sheet):
  - "Besonders gut (7-10)"  = count(rating >= 7)           [both polarities]
  - Positive criteria       : Positiv = count(rating > 0)  -> has good exposure
  - Negative criteria       : Positiv = count(rating > 9)  -> fully compliant
  - Negativ = group_size - Positiv
If a verification number is off, adjust FULL_COMPLIANCE_MIN / BESONDERS_GUT_MIN
or the POSITIVE_CRITERIA set below.
"""

import io
import os
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import requests
import streamlit as st

from llm import ask_question

st.set_page_config(page_title="CLEANVEST Fund Stats", layout="wide")

# --- Brand styling (turquoise & black) ---------------------------------
px.defaults.template = "plotly_white"
px.defaults.color_discrete_sequence = ["#0FB5A8", "#0E1B20", "#5BD1C6", "#6B7280", "#0A8C82"]

st.markdown(
    """
    <style>
      .block-container {padding-top: 2.2rem; max-width: 1200px;}
      .cv-header {background: linear-gradient(100deg,#0E1B20 0%,#0A8C82 62%,#0FB5A8 100%);
                  padding: 1.3rem 1.7rem; border-radius: 14px; margin-bottom: 1.3rem; color:#fff;}
      .cv-header h1 {margin:0; font-size:1.9rem; letter-spacing:.5px; color:#fff;}
      .cv-header p  {margin:.25rem 0 0; opacity:.85; font-size:.95rem;}
      [data-testid="stMetric"] {background:#F2FAF9; border:1px solid #D6EEEB;
                                border-radius:12px; padding:14px 16px;}
      [data-testid="stMetricValue"] {color:#0A8C82;}
      [data-testid="stMetricLabel"], [data-testid="stMetricLabel"] p {color:#0E1B20 !important;}
      button[data-baseweb="tab"] {font-size:1rem; font-weight:600;}
      .stTabs [aria-selected="true"] {color:#0FB5A8;}
      h2, h3 {color:#0E1B20;}
      /* callout notes */
      .cv-note {border-radius:10px; padding:.7rem 1rem; margin:.3rem 0 1.1rem;
                font-size:.92rem; line-height:1.45; border-left:5px solid;
                background:#F2FAF9; color:#0E1B20;}
      .cv-note .tag {display:inline-block; font-weight:700; font-size:.7rem;
                     letter-spacing:.07em; text-transform:uppercase; margin-right:.5rem;
                     color:#0A8C82;}
      .cv-note.info {border-color:#0FB5A8; background:#E8F8F6;}
      .cv-note.tip  {border-color:#0A8C82; background:#F2FAF9;}
      .cv-note.warn {border-color:#F9A825; background:#FFF6E6;}
      .cv-note.warn .tag {color:#B07400;}
      .cv-note.ok   {border-color:#0FB5A8; background:#E8F8F6;}
      .cv-note b {color:#0A8C82;}
    </style>
    """,
    unsafe_allow_html=True,
)

def note(text, kind="info", label=None):
    """Render a branded callout box. kind = info | tip | warn | ok."""
    labels = {"info": "Hinweis", "tip": "Tipp", "warn": "Achtung", "ok": "Kernaussagen"}
    lab = label or labels.get(kind, "Hinweis")
    st.markdown(
        f'<div class="cv-note {kind}"><span class="tag">{lab}</span>{text}</div>',
        unsafe_allow_html=True,
    )


def format_number_de(x, dec=2):
    """Deutsche Zahl mit fester Nachkommastellen-Anzahl: 1234.5 -> '1.234,50'."""
    try:
        xf = float(x)
    except (TypeError, ValueError):
        return x if x is not None else ""
    if pd.isna(xf):
        return ""
    return f"{xf:,.{dec}f}".replace(",", "#").replace(".", ",").replace("#", ".")


def format_number_smart_de(x):
    """Ganzzahlen ohne Nachkomma, sonst 2 Nachkommastellen."""
    try:
        xf = float(x)
    except (TypeError, ValueError):
        return x if x is not None else ""
    if pd.isna(xf):
        return ""
    return (f"{int(round(xf)):,}".replace(",", ".")) if xf.is_integer() else format_number_de(xf)


def format_percent_de(x):
    """Bruch (0.365) -> '36,50 %'."""
    try:
        xf = float(x)
    except (TypeError, ValueError):
        return x if x is not None else ""
    if pd.isna(xf):
        return ""
    return format_number_de(xf * 100) + " %"


def score_cell_style(v):
    """Ampel-Hintergrund für CLEANVEST Score: grün >=7, gelb 4-7, rot <=4."""
    try:
        x = float(v)
    except (TypeError, ValueError):
        return ""
    if pd.isna(x):
        return ""
    if x >= 7:
        return "background-color: #2e7d32; color: white; font-weight:600"
    if x > 4:
        return "background-color: #f9a825; color: #1A1A1A; font-weight:600"
    return "background-color: #c62828; color: white; font-weight:600"


def color_score_column(s):
    cell_styles = []                           # eine Farbregel je Zelle
    for cell_value in s:                       # jede Zelle der Spalte durchgehen
        cell_styles.append(score_cell_style(cell_value))
    return cell_styles


def status_cell_style(v):
    """Grün für OK (Haken), rot für Abweichung (Kreuz), sonst neutral."""
    s = str(v)
    if s == "✓":   # OK
        return "background-color: #2e7d32; color: white; font-weight:600"
    if s == "✗":   # Abweichung
        return "background-color: #c62828; color: white; font-weight:600"
    return ""


def color_status_column(col):
    cell_styles = []                           # eine Farbregel je Zelle
    for cell_value in col:                     # jede Zelle der Spalte durchgehen
        cell_styles.append(status_cell_style(cell_value))
    return cell_styles


def delta_cell_style(v):
    """Grün bei Verbesserung (>0), rot bei Verschlechterung (<0), sonst neutral."""
    try:
        x = float(v)
    except (TypeError, ValueError):
        return ""
    if pd.isna(x) or x == 0:
        return ""
    if x > 0:
        return "background-color: #2e7d32; color: white; font-weight:600"
    return "background-color: #c62828; color: white; font-weight:600"


def color_delta_column(s):
    cell_styles = []                           # eine Farbregel je Zelle
    for cell_value in s:                       # jede Zelle der Spalte durchgehen
        cell_styles.append(delta_cell_style(cell_value))
    return cell_styles


# --- Kategorien & Ampelfarben (einmal zentral, statt überall wiederholt) ---
CAT_COLORS = {
    "Nachhaltig": "#2e7d32",          # grün
    "Durchschnittlich": "#f9a825",    # gelb
    "Nicht nachhaltig": "#c62828",    # rot
}


def categorize(score):
    """Score -> Kategorie: >= 7 nachhaltig, <= 4 nicht nachhaltig, sonst durchschnittlich."""
    if score >= 7:
        return "Nachhaltig"
    if score <= 4:
        return "Nicht nachhaltig"
    return "Durchschnittlich"


# --- Configuration -----------------------------------------------------
RATING_PREFIX = "Rating: "
ANTEIL_PREFIX = "Anteil im Fonds: "
SCORE_COL = "CLEANVEST Score"
NAME_COL = "Fondsname"
ISIN_COL = "ISIN"
TYPE_COL = "Fondstype"
KAG_COL = "KAG"                             # Name der Fondsgesellschaft
DOMICILE_COL = "Fonds-Domizil"              # Land, in dem der Fonds aufgelegt ist
RATED_COL = "Ist bewertet"
VOLUME_EUR_COL = "Volumen - Betrag in EUR"
VOLUME_AMOUNT_COL = "Volumen - Betrag"      # Betrag in Originalwährung (vor EUR-Umrechnung)
VOLUME_CURRENCY_COL = "Volumen - Währung"   # Währungscode des Volumens, z. B. "USD"
VOLUME_DATE_COL = "Volumen - Datum"         # Stichtag, zu dem das Volumen gilt
FX_API_URL = "https://api.frankfurter.app"   # kostenlose EZB-Wechselkurse (Quelle der Originaldaten)

BESONDERS_GUT_MIN = 7      # rating >= 7  -> "Besonders gut (7-10)"
FULL_COMPLIANCE_MIN = 9    # rating  > 9  -> fully compliant (negative criteria)

POSITIVE_CRITERIA = {
    "Grüne Technologien",
    "Bildung & Gesundheit",
    "Commitment zu Emissionszielen",
    "Investitionen in Gesundheit",
    "Investitionen in Bildung",
    "Staaten mit hohem sozialen Fortschritt",
    "Staaten mit besonders niedriger Korruption",
}

SUST_LABELS = ["Nachhaltig (>= 7)", "Durchschnitt (4 < x < 7)", "Nicht nachhaltig (<= 4)"]

EXCLUDE_CRITERIA = []   # Kriterien hier eintragen, um sie überall auszublenden (leer = nichts ausblenden)

# Ordner, in dem die beiden CSV-Eingabedateien aus der Datenbank liegen.

app_folder = Path(__file__).parent.resolve()  # Ordner, in dem main.py liegt
INPUT_FOLDER = app_folder.joinpath("input")  # zentraler Ordner für CSV-Eingabedateien

# Die Dateinamen beginnen mit einem Datum, das sich jeden Monat aendert.
# Deshalb erkennen wir die Dateien nur am stabilen Teil ihres Namens.
UNIQUE_FUNDS_NAME_PART = "only_unique_funds"     # ein Datensatz pro Fonds (frueheres Blatt "DataUnique")
FUND_TRANCHES_NAME_PART = "all_fund_tranches"    # ein Datensatz pro Tranche (frueheres Blatt "DataAllTranches")


# ----------------------------------------------------------------------
# DATA LOADING
# ----------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def load_workbook(file_bytes):
    """Read DataUnique (funds) and DataAllTranches (fund_tranches)."""
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    unique_funds = pd.read_excel(xls, "DataUnique")
    fund_tranches = pd.read_excel(xls, "DataAllTranches") if "DataAllTranches" in xls.sheet_names else None
    return unique_funds, fund_tranches


def as_bool(series):
    """Coerce a 'rated' column to clean booleans regardless of how Excel stored it."""
    if series.dtype == bool:
        return series
    return series.astype(str).str.strip().str.lower().isin(["true", "wahr", "1", "ja", "yes"])


def find_input_file(folder_path, name_part):
    """Liefert den vollen Pfad der CSV-Datei in `folder_path`, deren Name `name_part` enthaelt."""
    # Gehe jeden Dateinamen im Ordner einzeln durch.
    for file_name in os.listdir(folder_path):
        # Wir suchen die CSV-Datei, deren Name den stabilen Teil enthaelt.
        if name_part in file_name and file_name.endswith(".csv"):
            # Baue den vollen Pfad (Ordner + Dateiname) und gib ihn zurueck.
            return os.path.join(folder_path, file_name)
    # Hierher kommen wir nur, wenn die erwartete Datei fehlt -> klar verstaendlicher Fehler.
    raise FileNotFoundError(f"Keine CSV-Datei mit '{name_part}' im Ordner {folder_path} gefunden.")


@st.cache_data(show_spinner=False)
def load_funds_and_tranches():
    """Liest die beiden CSV-Eingabedateien (Fonds + Tranchen) aus dem zentralen Eingabeordner."""
    # Finde die beiden Dateien anhand ihres stabilen Namensteils.
    unique_funds_path = find_input_file(INPUT_FOLDER, UNIQUE_FUNDS_NAME_PART)
    fund_tranches_path = find_input_file(INPUT_FOLDER, FUND_TRANCHES_NAME_PART)
    # Lies jede CSV-Datei in eine pandas-Tabelle (DataFrame) ein.
    unique_funds = pd.read_csv(unique_funds_path)
    fund_tranches = pd.read_csv(fund_tranches_path)
    # Gib beide Tabellen an den Aufrufer zurueck.
    return unique_funds, fund_tranches


def get_eur_rate(currency, date_text):
    """Holt den Wechselkurs '1 <currency> = ? EUR' fuer ein Datum von der EZB (Frankfurter API).

    Sonderfaelle (so wie in den urspruenglichen Kursdaten dokumentiert):
    - GBX sind britische Pence: 1 GBX = 0,01 GBP.
    - CNH (Offshore-Yuan) wird wie CNY behandelt.
    """
    if currency == "GBX":                                      # Pence -> erst GBP-Kurs, dann durch 100
        return get_eur_rate("GBP", date_text) * 0.01
    if currency == "CNH":                                      # Offshore-Yuan wie normalen Yuan behandeln
        currency = "CNY"

    url = f"{FX_API_URL}/{date_text}?from={currency}&to=EUR"   # Adresse fuer das Datum und die Waehrung
    response = requests.get(url, timeout=10)                   # Kurs von der EZB-API abfragen
    answer = response.json()                                   # Antwort als Daten (Dictionary) lesen
    if "rates" not in answer:                                  # Waehrung von der EZB nicht unterstuetzt?
        raise ValueError(f"Kein EUR-Kurs fuer Waehrung '{currency}' am {date_text} verfuegbar.")
    return answer["rates"]["EUR"]                              # den EUR-Kurs zurueckgeben


@st.cache_data(show_spinner=False)
def add_eur_volume_column(funds):
    """Berechnet je Fonds das Volumen in EUR und haengt es als neue Spalte an.

    Die Wechselkurse kommen vom Paket forex-python. Abgefragt werden nur Waehrung und Datum,
    niemals Fondsdaten. Gleiche Waehrung-Datum-Kombinationen werden nur einmal abgefragt.
    """
    funds = funds.copy()                          # Kopie, damit die Originaltabelle unveraendert bleibt

    # Die drei Volumen-Spalten kommen im Export zusaetzlich in Anfuehrungszeichen ("""..."""):
    # Entferne die Anfuehrungszeichen; der Betrag wird ausserdem in eine Zahl umgewandelt.
    funds[VOLUME_AMOUNT_COL] = funds[VOLUME_AMOUNT_COL].str.strip('"').astype(float)
    funds[VOLUME_CURRENCY_COL] = funds[VOLUME_CURRENCY_COL].str.strip('"')
    funds[VOLUME_DATE_COL] = funds[VOLUME_DATE_COL].str.strip('"')

    rate_by_currency_and_date = {}                # gemerkte Kurse: (Waehrung, Datum) -> Kurs nach EUR

    eur_amounts = []                              # EUR-Betrag je Fonds, in Zeilenreihenfolge
    for row_index in range(len(funds)):           # gehe jeden Fonds einzeln durch
        original_amount = funds[VOLUME_AMOUNT_COL].iloc[row_index]    # Betrag in Originalwaehrung (Zahl)
        currency = funds[VOLUME_CURRENCY_COL].iloc[row_index]         # Waehrungscode, z. B. "CHF"
        volume_date = funds[VOLUME_DATE_COL].iloc[row_index]          # Stichtag als Text "JJJJ-MM-TT"

        if currency == "EUR":                     # ein EUR-Fonds muss nicht umgerechnet werden
            eur_amounts.append(original_amount)   # Betrag unveraendert uebernehmen
            continue                              # weiter zum naechsten Fonds

        currency_and_date = (currency, volume_date)               # ein Kurs pro Waehrung und Datum
        if currency_and_date not in rate_by_currency_and_date:    # nur neue Kombinationen abfragen
            rate_by_currency_and_date[currency_and_date] = get_eur_rate(currency, volume_date)  # Kurs von der EZB holen

        eur_amounts.append(original_amount * rate_by_currency_and_date[currency_and_date])  # Betrag * Kurs

    funds[VOLUME_EUR_COL] = eur_amounts           # die fertige EUR-Spalte anhaengen
    return funds                                  # Tabelle mit neuer Spalte zurueckgeben


# ----------------------------------------------------------------------
# CALCULATIONS
# ----------------------------------------------------------------------
def is_positive(name):
    return name in POSITIVE_CRITERIA


def criteria_table(df_rated, rating_columns):
    """Per-criterion: Besonders gut / Positiv / Negativ counts, shares, Avg rating."""
    n = len(df_rated)
    rows = []
    for col in rating_columns:
        crit = col[len(RATING_PREFIX):]
        r = df_rated[col]
        besonders = int((r >= BESONDERS_GUT_MIN).sum())
        positiv = int((r > 0).sum()) if is_positive(crit) else int((r > FULL_COMPLIANCE_MIN).sum())
        negativ = n - positiv
        rows.append({
            "Kriterium": crit,
            "Besonders gut (7-10)": besonders,
            "Positiv": positiv,
            "Negativ": negativ,
            "Besonders gut %": besonders / n if n else np.nan,
            "Positiv %": positiv / n if n else np.nan,
            "Avg Rating": r.mean(),
        })
    return pd.DataFrame(rows).set_index("Kriterium")


def exposure_table(df_rated, exposure_columns):
    """Per-criterion exposure (Anteil im Fonds): Max / Min / Avg, in percent."""
    rows = []
    for col in exposure_columns:
        crit = col[len(ANTEIL_PREFIX):]
        a = df_rated[col]
        rows.append({"Kriterium": crit, "Max %": a.max(), "Min %": a.min(), "Avg %": a.mean()})
    return pd.DataFrame(rows).set_index("Kriterium")


def split_counts(unique_funds, fund_tranches, rated):
    """Headcount table: unique_funds funds + fund_tranches, split All / Rated / Not rated."""
    data = {"Unique Fonds": [len(unique_funds), int(rated.sum()), int((~rated).sum())]}
    tr = as_bool(fund_tranches[RATED_COL])
    data["Tranchen"] = [len(fund_tranches), int(tr.sum()), int((~tr).sum())]
    data["Avg Tranchen / Fonds"] = [
        len(fund_tranches) / len(unique_funds) if len(unique_funds) else np.nan,
        tr.sum() / rated.sum() if rated.sum() else np.nan,
        (~tr).sum() / (~rated).sum() if (~rated).sum() else np.nan,
    ]
    return pd.DataFrame(data, index=["Alle", "Bewertete Fonds", "Nicht bewertete Fonds"]).T


def kag_statistics(funds):
    """Berechnet je KAG (Fondsgesellschaft): Anzahl Fonds, Durchschnitts-Score und Volumen in Mio. EUR."""
    statistics_rows = []                               # sammelt eine Ergebniszeile pro KAG
    for kag_name in funds[KAG_COL].unique():           # gehe jede vorkommende KAG einzeln durch
        funds_of_this_kag = funds[funds[KAG_COL] == kag_name]   # nur die Fonds dieser KAG
        statistics_rows.append({                       # eine Ergebniszeile fuer diese KAG
            "KAG": kag_name,                                                  # Name der Fondsgesellschaft
            "Anzahl Fonds": len(funds_of_this_kag),                          # Anzahl der Fonds dieser KAG
            "Ø Score": funds_of_this_kag[SCORE_COL].mean(),                  # durchschnittlicher CLEANVEST Score
            "Volumen Mio. EUR": funds_of_this_kag[VOLUME_EUR_COL].sum() / 1e6,  # Gesamtvolumen in Mio. EUR
        })
    statistics_table = pd.DataFrame(statistics_rows)   # die gesammelten Zeilen werden zu einer Tabelle
    statistics_table = statistics_table.sort_values("Anzahl Fonds", ascending=False)  # groesste KAG zuerst
    return statistics_table


def domicile_statistics(funds):
    """Berechnet je Land (Fonds-Domizil): Anzahl Fonds und Durchschnitts-Score."""
    statistics_rows = []                               # sammelt eine Ergebniszeile pro Land
    for country_name in funds[DOMICILE_COL].unique():  # gehe jedes vorkommende Land einzeln durch
        funds_of_this_country = funds[funds[DOMICILE_COL] == country_name]   # nur die Fonds dieses Landes
        statistics_rows.append({                       # eine Ergebniszeile fuer dieses Land
            "Domizil": country_name,                                 # Name des Landes
            "Anzahl Fonds": len(funds_of_this_country),              # Anzahl der Fonds dieses Landes
            "Ø Score": funds_of_this_country[SCORE_COL].mean(),      # durchschnittlicher CLEANVEST Score
        })
    statistics_table = pd.DataFrame(statistics_rows)   # die gesammelten Zeilen werden zu einer Tabelle
    statistics_table = statistics_table.sort_values("Anzahl Fonds", ascending=False)  # groesstes Land zuerst
    return statistics_table


# ----------------------------------------------------------------------
# UI
# ----------------------------------------------------------------------
st.markdown(
    """
    <div class="cv-header">
      <h1>CLEANVEST</h1>
      <p>Fund Statistics Dashboard</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# Beim Start die beiden CSV-Eingabedateien (Fonds + Tranchen) aus dem zentralen Ordner laden.
unique_funds, fund_tranches = load_funds_and_tranches()

# Berechne die Spalte "Volumen - Betrag in EUR" mit dem Paket forex-python (kommt nicht mehr aus der Datei).
unique_funds = add_eur_volume_column(unique_funds)

is_rated = as_bool(unique_funds[RATED_COL])
rating_columns = []                                # alle Rating-Spalten sammeln
for column_name in unique_funds.columns:           # jede Spalte durchgehen
    if column_name.startswith(RATING_PREFIX):      # nur "Rating: ..."-Spalten behalten
        rating_columns.append(column_name)
exposure_columns = []                              # alle Anteil-Spalten sammeln
for column_name in unique_funds.columns:           # jede Spalte durchgehen
    if column_name.startswith(ANTEIL_PREFIX):      # nur "Anteil im Fonds: ..."-Spalten behalten
        exposure_columns.append(column_name)
rated_funds = unique_funds[is_rated]

extract_date = pd.to_datetime(unique_funds["Datum des Auszugs"]).max().strftime("%d.%m.%Y")

st.caption(f"Datenabzug: {extract_date} · {len(unique_funds)} Fonds · {len(rating_columns)} Kriterien")

tab_overview, tab_dist, tab_criteria, tab_extra, tab_kag, tab_domicile, tab_compare, tab_month, tab_search, tab_fx, tab_verify, tab_chat = st.tabs(
    ["Überblick", "Score-Verteilung", "Kriterien", "Weitere Analysen", "KAG", "Domizil",
     "Fonds-Vergleich", "Monatsvergleich", "Fonds-Suche", "Volumen → EUR", "Verifikation", "KI-Assistent"]
)

# --- OVERVIEW ----------------------------------------------------------
with tab_overview:
    note("Eckdaten des Datenabzugs: Anzahl Fonds und Tranchen, Aufteilung nach Typ, "
         "Nachhaltigkeits-Einteilung und Volumen.", "info")

    c1, c2, c3 = st.columns(3)
    c1.metric("Unique Fonds", len(unique_funds))
    c2.metric("Bewertet", int(is_rated.sum()))
    c3.metric("Nicht bewertet", int((~is_rated).sum()))

    if len(rated_funds):
        nach = int((rated_funds[SCORE_COL] >= 7).sum())
        share = nach / len(rated_funds) * 100
        top = rated_funds.loc[rated_funds[SCORE_COL].idxmax()]
        bits = [f"<b>{share:.0f}%</b> der bewerteten Fonds sind nachhaltig (Score &ge; 7)",
                f"Ø CLEANVEST Score: <b>{rated_funds[SCORE_COL].mean():.2f}</b>",
                f"Bester Fonds: <b>{top[NAME_COL]}</b> ({top[SCORE_COL]:.1f})",
                f"Volumen (bewertet): <b>{rated_funds[VOLUME_EUR_COL].sum() / 1e9:.0f} Mrd. €</b>"]
        note(" &nbsp;·&nbsp; ".join(bits), "ok")

    st.subheader("Fonds & Tranchen")
    st.dataframe(split_counts(unique_funds, fund_tranches, is_rated).style.format(format_number_smart_de), use_container_width=True)

    st.subheader("Unique Fonds je Typ")
    by_type = pd.DataFrame({
        "Alle": unique_funds.groupby(TYPE_COL).size(),
        "Bewertete Fonds": rated_funds.groupby(TYPE_COL).size(),
        "Nicht bewertete Fonds": unique_funds[~is_rated].groupby(TYPE_COL).size(),
    }).fillna(0).astype(int)
    by_type.loc["Summe"] = by_type.sum()
    st.dataframe(by_type, use_container_width=True)

    st.subheader("Nachhaltigkeits-Einteilung (bewertete Fonds)")
    s = rated_funds[SCORE_COL]
    sust = pd.DataFrame(
        {"Fonds": [int((s >= 7).sum()), int(((s > 4) & (s < 7)).sum()), int((s <= 4).sum())]},
        index=SUST_LABELS,
    )
    sust["Anteil %"] = (sust["Fonds"] / sust["Fonds"].sum() * 100).round(2)

    def color_sustainability_row(row):
        name = str(row.name)
        if name.startswith("Nachhaltig"):
            bg, fg = "#2e7d32", "white"          # grün
        elif name.startswith("Durchschnitt"):
            bg, fg = "#f9a825", "#1A1A1A"         # gelb
        else:
            bg, fg = "#c62828", "white"          # rot
        return [f"background-color: {bg}; color: {fg}; font-weight:600"] * len(row)

    st.dataframe(sust.style.apply(color_sustainability_row, axis=1).format({"Anteil %": format_number_de}), use_container_width=True)

    st.subheader("Volumen (EUR)")
    vol = pd.DataFrame({"Volumen EUR": [
        unique_funds[VOLUME_EUR_COL].sum(),
        rated_funds[VOLUME_EUR_COL].sum(),
        unique_funds[~is_rated][VOLUME_EUR_COL].sum(),
    ]}, index=["Alle", "Bewertete Fonds", "Nicht bewertete Fonds"])
    vol["Volumen Mrd. EUR"] = vol["Volumen EUR"] / 1e9

    def format_eur_de(x):
        # 769106791785.40 -> '769.106.791.785,40'
        return f"{x:,.2f}".replace(",", "#").replace(".", ",").replace("#", ".")

    st.dataframe(
        vol.style.format({"Volumen EUR": format_eur_de, "Volumen Mrd. EUR": "{:.2f}"}),
        use_container_width=True,
    )

# --- SCORE DISTRIBUTION (core deliverable) -----------------------------
with tab_dist:
    st.subheader("Score-Verteilung der bewerteten Fonds")
    note("Verteilung des CLEANVEST Scores. <b>Grün</b> = nachhaltig (&ge; 7), "
         "<b>gelb</b> = durchschnittlich (4–7), <b>rot</b> = nicht nachhaltig (&le; 4).", "tip")
    bins = list(range(0, 11))
    labels = []                                    # Beschriftungen der Score-Bereiche: "0-1", "1-2", ...
    for lower_bound in range(10):
        labels.append(f"{lower_bound}-{lower_bound + 1}")
    cats = pd.cut(rated_funds[SCORE_COL], bins=bins, labels=labels, include_lowest=True)
    dist = cats.value_counts().reindex(labels, fill_value=0).reset_index()
    dist.columns = ["Score-Bereich", "Fonds"]
    dist["Anteil %"] = (dist["Fonds"] / dist["Fonds"].sum() * 100).round(2)

    def band(x):
        lo = int(x.split("-")[0])
        return "Nachhaltig" if lo >= 7 else ("Nicht nachhaltig" if lo < 4 else "Durchschnittlich")
    dist["Kategorie"] = dist["Score-Bereich"].map(band)

    fig = px.bar(dist, x="Score-Bereich", y="Fonds", color="Kategorie", text="Fonds",
                 title="Verteilung CLEANVEST Score",
                 color_discrete_map=CAT_COLORS)
    fig.update_layout(font=dict(size=15), title_font_size=22,
                      xaxis_title_font_size=15, yaxis_title_font_size=15,
                      legend_font_size=14, legend_title_font_size=14)
    fig.update_traces(textfont_size=15, textposition="outside")
    fig.update_xaxes(tickfont_size=14)
    fig.update_yaxes(tickfont_size=13)
    st.plotly_chart(fig, use_container_width=True)
    dist_show = dist.set_index("Score-Bereich")

    def color_category_row(row):
        cat = str(row["Kategorie"])
        if cat == "Nachhaltig":
            bg, fg = "#2e7d32", "white"           # grün
        elif cat == "Durchschnittlich":
            bg, fg = "#f9a825", "#1A1A1A"          # gelb
        else:
            bg, fg = "#c62828", "white"           # rot
        return [f"background-color: {bg}; color: {fg}; font-weight:600"] * len(row)

    st.dataframe(
        dist_show.style.apply(color_category_row, axis=1).format({"Anteil %": format_number_de}),
        use_container_width=True,
    )

    a, b, c = st.columns(3)
    a.metric("Ø Score", format_number_de(rated_funds[SCORE_COL].mean()))
    b.metric("Median", format_number_de(rated_funds[SCORE_COL].median()))
    c.metric("Std", format_number_de(rated_funds[SCORE_COL].std()))

# --- CRITERIA ----------------------------------------------------------
with tab_criteria:
    st.subheader("Kriterien-Statistik")
    note("Pro Kriterium: Anteil positiver/negativer Fonds und Ø Rating. Wähle einen "
         "Subtyp, um ihn mit allen Fonds zu vergleichen (Spalte Differenz).", "tip")
    types = ["Alle Fondstypen"] + (sorted(rated_funds[TYPE_COL].dropna().unique())
                                   if TYPE_COL in rated_funds.columns else [])
    sub = st.selectbox("Subtyp zum Vergleich", types)

    all_tbl = criteria_table(rated_funds, rating_columns)
    if sub == "Alle Fondstypen":
        st.dataframe(all_tbl.style.format({"Avg Rating": format_number_de, "Besonders gut %": format_percent_de, "Positiv %": format_percent_de}), use_container_width=True)
    else:
        sub_df = rated_funds[rated_funds[TYPE_COL] == sub]
        sub_tbl = criteria_table(sub_df, rating_columns)
        cmp = pd.DataFrame({
            "Avg Rating (Alle)": all_tbl["Avg Rating"],
            f"Avg Rating ({sub})": sub_tbl["Avg Rating"],
        })
        cmp["Differenz"] = cmp.iloc[:, 1] - cmp.iloc[:, 0]
        cmp["Positiv % (Alle)"] = all_tbl["Positiv %"]
        cmp[f"Positiv % ({sub})"] = sub_tbl["Positiv %"]
        _cmp_fmt = {"Avg Rating (Alle)": format_number_de, f"Avg Rating ({sub})": format_number_de, "Differenz": format_number_de, "Positiv % (Alle)": format_percent_de, f"Positiv % ({sub})": format_percent_de}
        st.dataframe(cmp.style.format(_cmp_fmt), use_container_width=True)

    st.subheader("Anteile im Fonds (Exposure, %)")
    exp_df = rated_funds if sub == "Alle Fondstypen" else rated_funds[rated_funds[TYPE_COL] == sub]
    st.dataframe(exposure_table(exp_df, exposure_columns).style.format({"Max %": format_number_de, "Min %": format_number_de, "Avg %": format_number_de}), use_container_width=True)

# --- EXTRA ANALYSES ----------------------------------------------------
with tab_extra:
    note("Zusätzliche Auswertungen: beste und schlechteste Fonds, Ø Score je Typ, "
         "Verteilung der Fonds und Aufschlüsselung nach KAG/Domizil.", "tip")

    st.subheader("Top 10 / Schlechteste 10 Fonds (CLEANVEST Score)")
    show = [NAME_COL, ISIN_COL, TYPE_COL, SCORE_COL]   # anzuzeigende Spalten, immer in dieser Reihenfolge
    left, right = st.columns(2)
    left.markdown("**Top 10**")
    top10 = rated_funds.nlargest(10, SCORE_COL)[show].reset_index(drop=True)
    top10.index = range(1, len(top10) + 1)
    top10.index.name = "Rang"
    left.dataframe(top10.style.format({SCORE_COL: format_number_de}).apply(color_score_column, subset=[SCORE_COL]), use_container_width=True)
    right.markdown("**Schlechteste 10**")
    worst10 = rated_funds.nsmallest(10, SCORE_COL)[show].reset_index(drop=True)
    worst10.index = range(1, len(worst10) + 1)
    worst10.index.name = "Rang"
    right.dataframe(worst10.style.format({SCORE_COL: format_number_de}).apply(color_score_column, subset=[SCORE_COL]), use_container_width=True)

    st.subheader("Ø CLEANVEST Score je Fondstyp")
    avg_type = (rated_funds.groupby(TYPE_COL)[SCORE_COL].mean().round(2)
                .sort_values(ascending=False).reset_index())
    avg_type["Kategorie"] = avg_type[SCORE_COL].apply(categorize)
    st.plotly_chart(
        px.bar(avg_type, x=TYPE_COL, y=SCORE_COL, color="Kategorie",
               title="Ø Score je Typ", range_y=[0, 10], color_discrete_map=CAT_COLORS),
        use_container_width=True,
    )

    st.subheader("Verteilung der Fonds nach Typ")
    st.plotly_chart(px.pie(unique_funds, names=TYPE_COL, title="Fonds nach Typ"), use_container_width=True)

    st.subheader("Ø Score nach Dimension")
    dim = st.selectbox("Dimension wählen", [KAG_COL, DOMICILE_COL])
    avg_dim = (rated_funds.groupby(dim)[SCORE_COL].mean().round(2)
               .sort_values(ascending=False).head(20).reset_index())
    avg_dim["Kategorie"] = avg_dim[SCORE_COL].apply(categorize)
    st.plotly_chart(
        px.bar(avg_dim, x=dim, y=SCORE_COL, color="Kategorie",
               title=f"Ø Score nach {dim}", range_y=[0, 10], color_discrete_map=CAT_COLORS),
        use_container_width=True,
    )

# --- FUND COMPARE ------------------------------------------------------
with tab_compare:
    st.subheader("Fonds-Vergleich")
    note("Wähle zwei oder mehr Fonds, um ihre CLEANVEST Scores und Kriterien-Ratings "
         "direkt zu vergleichen.", "tip")
    options = sorted(rated_funds[NAME_COL].dropna().astype(str).unique().tolist())
    chosen = st.multiselect("Fonds auswählen (mindestens 2)", options)
    if len(chosen) >= 2:
        selected_funds = (rated_funds[rated_funds[NAME_COL].astype(str).isin(chosen)]
               .drop_duplicates(subset=[NAME_COL]).set_index(NAME_COL))
        st.markdown("**CLEANVEST Score**")
        st.dataframe(selected_funds[[SCORE_COL]].T.style.format(format_number_de).apply(color_score_column), use_container_width=True)
        total_score_compare_df = selected_funds.reset_index()
        total_score_compare_df["Kategorie"] = total_score_compare_df[SCORE_COL].apply(categorize)
        st.plotly_chart(
            px.bar(total_score_compare_df, x=NAME_COL, y=SCORE_COL, color="Kategorie", text=SCORE_COL,
                   title="CLEANVEST Score im Vergleich", range_y=[0, 10],
                   color_discrete_map=CAT_COLORS),
            use_container_width=True,
        )
        criteria_columns = []                                 # Rating-Spalten der ausgewaehlten Fonds
        for column_name in selected_funds.columns:
            if column_name.startswith(RATING_PREFIX):
                criteria_columns.append(column_name)
        if criteria_columns:
            criteria_comparison_table = selected_funds[criteria_columns].T
            criteria_comparison_table.index = criteria_comparison_table.index.str.removeprefix(RATING_PREFIX)
            criteria_comparison_table.index.name = "Kriterium"
            st.markdown("**Ratings je Kriterium**")
            st.dataframe(criteria_comparison_table.style.format(format_number_de).apply(color_score_column), use_container_width=True)
    else:
        st.caption("Bitte mindestens zwei Fonds auswählen.")

# --- MONTH-OVER-MONTH (Monatsvergleich) --------------------------------
with tab_month:
    st.subheader("Monatsvergleich")
    note("Lade die Datei des Vormonats hoch. Die App vergleicht die CLEANVEST Scores Fonds für Fonds: "
         "wer hat sich verbessert oder verschlechtert, wer überschreitet die Nachhaltigkeits-Schwelle (7) "
         "und welche Fonds sind neu bzw. weggefallen.", "tip")

    prev_file = st.file_uploader("Vormonats-Datei hochladen (Excel)",
                                 type=["xlsx", "xls"], key="prev_month")
    if prev_file is None:
        st.caption("Bitte die Vormonats-Datei hochladen, um den Vergleich zu starten.")
    else:
        prev_unique, _ = load_workbook(prev_file.getvalue())
        if ISIN_COL not in prev_unique.columns or SCORE_COL not in prev_unique.columns:
            st.error(f"In der Vormonats-Datei fehlt '{ISIN_COL}' oder '{SCORE_COL}'.")
        else:
            current = (rated_funds[[ISIN_COL, SCORE_COL]].dropna(subset=[ISIN_COL])
                   .drop_duplicates(ISIN_COL).set_index(ISIN_COL)[SCORE_COL])
            prev_rated = (prev_unique[as_bool(prev_unique[RATED_COL])]
                          if RATED_COL in prev_unique.columns else prev_unique)
            prev = (prev_rated[[ISIN_COL, SCORE_COL]].dropna(subset=[ISIN_COL])
                    .drop_duplicates(ISIN_COL).set_index(ISIN_COL)[SCORE_COL])

            common = current.index.intersection(prev.index)
            if len(common) == 0:
                st.warning("Keine gemeinsamen Fonds gefunden — stimmt die Schlüsselspalte überein?")
            else:
                comp = pd.DataFrame({"Vormonat": prev.reindex(common),
                                     "Aktuell": current.reindex(common)})
                comp["Differenz"] = (comp["Aktuell"] - comp["Vormonat"]).round(2)

                new_funds = current.index.difference(prev.index)
                gone_funds = prev.index.difference(current.index)

                became_sustainable = int(((comp["Vormonat"] < 7) & (comp["Aktuell"] >= 7)).sum())
                not_sustainable_anymore = int(((comp["Vormonat"] >= 7) & (comp["Aktuell"] < 7)).sum())

                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Verbessert", int((comp["Differenz"] > 0).sum()))
                m2.metric("Verschlechtert", int((comp["Differenz"] < 0).sum()))
                m3.metric("Neu nachhaltig", became_sustainable)
                m4.metric("Nicht mehr nachhaltig", not_sustainable_anymore)

                d1, d2, d3 = st.columns(3)
                d1.metric("Ø Differenz", format_number_de(comp["Differenz"].mean()))
                d2.metric("Neue Fonds", len(new_funds))
                d3.metric("Weggefallene Fonds", len(gone_funds))

                st.subheader("Größte Veränderungen")
                changed = comp[comp["Differenz"] != 0].sort_values("Differenz", key=lambda s: s.abs(), ascending=False)
                st.dataframe(
                    changed.style.format({"Vormonat": format_number_de, "Aktuell": format_number_de, "Differenz": format_number_de})
                           .apply(color_score_column, subset=["Vormonat", "Aktuell"])
                           .apply(color_delta_column, subset=["Differenz"]),
                    use_container_width=True,
                )

                if len(new_funds):
                    st.markdown("**Neue Fonds (nur aktuell)**")
                    st.dataframe(
                        current.reindex(new_funds).rename("Aktueller Score").to_frame()
                           .style.format(format_number_de).apply(color_score_column),
                        use_container_width=True)
                if len(gone_funds):
                    st.markdown("**Weggefallene Fonds (nur Vormonat)**")
                    st.dataframe(
                        prev.reindex(gone_funds).rename("Score Vormonat").to_frame()
                            .style.format(format_number_de).apply(color_score_column),
                        use_container_width=True)


# --- FUND SEARCH (ISIN) ------------------------------------------------
with tab_search:
    st.subheader("Fonds-Suche (ISIN)")
    note("Gib eine ISIN (oder einen Teil davon) ein, um den Fonds mit allen Details und seinem "
         "Kriterien-Profil zu sehen.", "tip")
    q = st.text_input("ISIN eingeben").strip()
    if not q:
        st.caption("Bitte eine ISIN eingeben.")
    else:
        matches = unique_funds[unique_funds[ISIN_COL].astype(str).str.contains(q, case=False, na=False)]
        if matches.empty:
            st.warning("Kein Fonds mit dieser ISIN gefunden.")
        else:
            if len(matches) > 1:
                labels = [f"{r[ISIN_COL]}  —  {r[NAME_COL]}" for _, r in matches.iterrows()]
                pick = st.selectbox(f"{len(matches)} Treffer — Fonds wählen", labels)
                row = matches.iloc[labels.index(pick)]
            else:
                row = matches.iloc[0]

            st.markdown(f"### {row[NAME_COL]}")
            k1, k2, k3 = st.columns(3)
            sc = row[SCORE_COL]
            k1.metric("CLEANVEST Score", format_number_de(sc) if pd.notna(sc) else "nicht bewertet")
            k2.metric("Fondstyp", str(row[TYPE_COL]))
            k3.metric("Volumen", f"{row[VOLUME_EUR_COL] / 1e6:.1f} Mio. €")
            st.caption(f"ISIN: {row[ISIN_COL]}")

            rcols = []                                 # alle Rating-Spalten
            for column_name in unique_funds.columns:
                if column_name.startswith(RATING_PREFIX):
                    rcols.append(column_name)
            rating_values = []                     # Bewertung dieses Fonds je Kriterium
            criterion_names = []                   # Kriteriumsname ohne "Rating: "-Praefix
            for column_name in rcols:
                rating_values.append(row[column_name])
                criterion_names.append(column_name[len(RATING_PREFIX):])
            rat = pd.DataFrame({"Rating": rating_values}, index=criterion_names)
            rat.index.name = "Kriterium"
            st.markdown("**Kriterien-Ratings**")
            st.dataframe(rat.style.format(format_number_de).apply(color_score_column), use_container_width=True)

            radar_df = rat.reset_index().dropna(subset=["Rating"])
            if len(radar_df) >= 3:
                radar = px.line_polar(radar_df, r="Rating", theta="Kriterium",
                                      line_close=True, range_r=[0, 10], title="Kriterien-Profil")
                radar.update_traces(fill="toself", line_color="#0A8C82")
                st.plotly_chart(radar, use_container_width=True)


# --- VOLUMEN IN EUR (EZB / Frankfurter-API) ----------------------------
with tab_fx:
    st.subheader("Volumen in EUR")
    note("Die Fonds-Volumina werden mit aktuellen EZB-Wechselkursen (Frankfurter-API) in EUR "
         "umgerechnet - nach Waehrung und Datum. Es werden keine Fondsdaten gesendet.", "tip")
    # Zeige pro Fonds: Name, Waehrung, Originalbetrag und den umgerechneten EUR-Betrag.
    columns_to_show = [NAME_COL, VOLUME_CURRENCY_COL, VOLUME_AMOUNT_COL, VOLUME_EUR_COL]
    # Betrag und EUR-Betrag in deutscher Schreibweise formatieren (1.234,50).
    st.dataframe(
        unique_funds[columns_to_show].style.format({VOLUME_AMOUNT_COL: format_number_de, VOLUME_EUR_COL: format_number_de}),
        use_container_width=True,
    )
    # Gesamtsumme aller Fonds-Volumina in EUR (in Milliarden, deutsch formatiert).
    total_volume_eur = unique_funds[VOLUME_EUR_COL].sum()
    st.metric("Gesamtvolumen", f"{format_number_de(total_volume_eur / 1e9, 2)} Mrd. EUR")


# --- KAG-STATISTIK -----------------------------------------------------
with tab_kag:
    st.subheader("KAG-Statistik")
    note("Auswertung je KAG (Fondsgesellschaft): Anzahl Fonds, durchschnittlicher Score und "
         "Gesamtvolumen. Basis sind die bewerteten Fonds.", "info")
    # Berechne die Statistik je KAG aus den bewerteten Fonds.
    kag_table = kag_statistics(rated_funds)
    # Zeige die Tabelle mit deutscher Zahlenformatierung; die Score-Spalte wird per Ampel eingefaerbt.
    st.dataframe(
        kag_table.style.format({"Ø Score": format_number_de, "Volumen Mio. EUR": format_number_de})
                 .apply(color_score_column, subset=["Ø Score"]),
        use_container_width=True,
    )
    # Balkendiagramm: die 15 KAGs mit den meisten Fonds.
    top_kags = kag_table.head(15)
    st.plotly_chart(
        px.bar(top_kags, x="KAG", y="Anzahl Fonds", title="KAGs mit den meisten Fonds"),
        use_container_width=True,
    )


# --- DOMIZIL-STATISTIK + KARTE -----------------------------------------
with tab_domicile:
    st.subheader("Domizil-Statistik")
    note("Auswertung je Land (Fonds-Domizil): Anzahl Fonds und durchschnittlicher Score. "
         "Basis sind die bewerteten Fonds.", "info")
    # Berechne die Statistik je Land aus den bewerteten Fonds.
    domicile_table = domicile_statistics(rated_funds)
    # Zeige die Tabelle mit deutscher Zahlenformatierung; die Score-Spalte wird per Ampel eingefaerbt.
    st.dataframe(
        domicile_table.style.format({"Ø Score": format_number_de})
                      .apply(color_score_column, subset=["Ø Score"]),
        use_container_width=True,
    )
    # Weltkarte: jedes Land nach Anzahl Fonds eingefaerbt (Plotly, ohne Zusatzpakete).
    map_figure = px.choropleth(
        domicile_table,
        locations="Domizil",                  # die Werte in dieser Spalte sind Laendernamen
        locationmode="country names",         # Plotly erkennt englische Laendernamen automatisch
        color="Anzahl Fonds",                 # Farbe richtet sich nach der Anzahl Fonds
        hover_data=["Ø Score"],               # beim Drueberfahren auch den Score zeigen
        title="Anzahl Fonds je Land",
        color_continuous_scale="Tealgrn",     # gruen-tuerkise Farbskala (passt zum Look)
    )
    st.plotly_chart(map_figure, use_container_width=True)


# --- VERIFICATION -- pausiert: das Stats-Blatt gibt es im CSV-Datensatz nicht mehr ---
with tab_verify:
    st.subheader("Verifikation")
    # Frueher wurden die App-Werte mit dem Excel-Blatt "Stats" verglichen.
    # Der neue CSV-Datensatz enthaelt dieses Blatt nicht mehr, daher pausiert der Abgleich.
    note("Die Verifikation verglich die App-Werte mit dem Excel-Blatt „Stats“. Der neue "
         "CSV-Datensatz enthaelt dieses Blatt nicht mehr, daher ist dieser Abgleich aktuell "
         "nicht verfuegbar.", "warn")


# --- KI-ASSISTENT (Ollama oder lokales transformers-Modell) ------------
with tab_chat:
    st.subheader("KI-Assistent")
    st.caption("Lokaler Chatbot - läuft komplett offline auf deinem Rechner. "
               "Die Fonds-Daten verlassen den Laptop nicht.")

    OLLAMA_URL = "http://localhost:11434/api/chat"

    @st.cache_data(ttl=15, show_spinner=False)
    def _ollama_status():
        import requests
        try:
            requests.get("http://localhost:11434/api/tags", timeout=1.5)
            return True
        except Exception:
            return False

    with st.expander("Einstellungen"):
        backend = st.radio("Backend", ["Lokal (transformers)", "Ollama"], horizontal=True)
        if backend == "Ollama":
            chat_model = st.text_input("Modell (in Ollama installiert)", value="llama3.2:3b")
        else:
            st.caption("Modell: Qwen/Qwen2.5-3B-Instruct (wird beim ersten Aufruf geladen; "
                       "die Oberfläche pausiert währenddessen).")
        if st.button("Verlauf löschen"):
            st.session_state.pop("chat_history", None)
            st.rerun()

    if backend == "Ollama":
        if _ollama_status():
            st.markdown(
                f"<span style='color:#2e7d32;font-weight:600'>&#9679; Ollama verbunden</span>"
                f"<span style='color:#6B7280'> &nbsp;·&nbsp; Modell {chat_model}</span>",
                unsafe_allow_html=True)
        else:
            st.markdown(
                "<span style='color:#c62828;font-weight:600'>&#9679; Ollama nicht erreichbar</span>"
                "<span style='color:#6B7280'> &nbsp;·&nbsp; Ollama-App starten, dann oben auf Reload</span>",
                unsafe_allow_html=True)

    def _build_context():
        """Kompakte Daten-Zusammenfassung als Grundlage für den Assistenten."""
        lines = [f"Anzahl Fonds gesamt: {len(unique_funds)}, davon bewertet: {int(is_rated.sum())}."]
        if len(rated_funds):
            nach = int((rated_funds[SCORE_COL] >= 7).sum())
            lines.append(f"Nachhaltig (Score >= 7): {nach} ({nach / len(rated_funds) * 100:.1f} %). "
                         f"Durchschnittlicher CLEANVEST Score: {rated_funds[SCORE_COL].mean():.2f}.")
            top = rated_funds.nlargest(5, SCORE_COL)[[NAME_COL, SCORE_COL]]
            worst = rated_funds.nsmallest(5, SCORE_COL)[[NAME_COL, SCORE_COL]]
            top_parts = []                             # je Top-Fonds ein Text "Name (Score)"
            for _, fund_row in top.iterrows():
                top_parts.append(f"{fund_row[NAME_COL]} ({fund_row[SCORE_COL]:.1f})")
            lines.append("Top 5 Fonds (Score): " + "; ".join(top_parts))
            worst_parts = []                           # je schwaechsten Fonds ein Text "Name (Score)"
            for _, fund_row in worst.iterrows():
                worst_parts.append(f"{fund_row[NAME_COL]} ({fund_row[SCORE_COL]:.1f})")
            lines.append("Schlechteste 5 Fonds (Score): " + "; ".join(worst_parts))
        ct = criteria_table(rated_funds, rating_columns)["Avg Rating"]
        strongest_parts = []                           # je Kriterium ein Text "Name: Ø Rating"
        for criterion_name, avg_rating in ct.sort_values(ascending=False).head(5).items():
            strongest_parts.append(f"{criterion_name}: {avg_rating:.2f}")
        lines.append("Stärkste Kriterien (Ø Rating): " + "; ".join(strongest_parts))
        weakest_parts = []                             # je Kriterium ein Text "Name: Ø Rating"
        for criterion_name, avg_rating in ct.sort_values().head(5).items():
            weakest_parts.append(f"{criterion_name}: {avg_rating:.2f}")
        lines.append("Schwächste Kriterien (Ø Rating): " + "; ".join(weakest_parts))
        return "\n".join(lines)

    if "chat_history" not in st.session_state:
        st.session_state.chat_history = []

    # Beispiel-Fragen anzeigen, solange der Verlauf leer ist
    pending = None
    if not st.session_state.chat_history:
        st.write("")
        st.markdown("<span style='color:#6B7280'>Beispiele zum Starten:</span>", unsafe_allow_html=True)
        examples = ["Welcher Fonds hat den höchsten Score?",
                    "Bei welchem Kriterium schneiden die Fonds am schlechtesten ab?",
                    "Wie viele Fonds sind nachhaltig?"]
        for col, q in zip(st.columns(len(examples)), examples):
            if col.button(q, use_container_width=True):
                pending = q

    AVATAR = {"user": "🧑", "assistant": "🤖"}
    for m in st.session_state.chat_history:
        with st.chat_message(m["role"], avatar=AVATAR[m["role"]]):
            st.markdown(m["content"])

    prompt = st.chat_input("Frag etwas zu den Fonds ...") or pending
    if prompt:
        st.session_state.chat_history.append({"role": "user", "content": prompt})
        with st.chat_message("user", avatar=AVATAR["user"]):
            st.markdown(prompt)

        system = ("Du bist ein hilfreicher Assistent für eine Fonds-Nachhaltigkeits-App (CLEANVEST). "
                  "Antworte auf Deutsch, kurz und sachlich. Stütze dich auf die folgenden Daten; "
                  "wenn etwas nicht in den Daten steht, sage das ehrlich.\n\nDATEN:\n" + _build_context())
        messages = [{"role": "system", "content": system}] + st.session_state.chat_history

        def stream_answer():
            import json

            import requests
            try:
                with requests.post(OLLAMA_URL, json={"model": chat_model, "messages": messages,
                                                     "stream": True}, stream=True, timeout=300) as resp:
                    resp.raise_for_status()
                    for line in resp.iter_lines():
                        if not line:
                            continue
                        data = json.loads(line.decode("utf-8"))
                        chunk = data.get("message", {}).get("content", "")
                        if chunk:
                            yield chunk
                        if data.get("done"):
                            break
            except requests.exceptions.ConnectionError:
                yield ("Ollama ist nicht erreichbar. Bitte Ollama starten und das Modell laden: "
                       f"im Terminal `ollama pull {chat_model}` ausführen. (Adresse: http://localhost:11434)")
            except Exception as e:
                yield f"Fehler beim Abruf des lokalen Modells: {e}"

        with st.chat_message("assistant", avatar=AVATAR["assistant"]):
            if backend == "Ollama":
                answer = st.write_stream(stream_answer)
            else:
                with st.spinner("Antwort wird lokal generiert ..."):
                    answer = ask_question(messages)
                st.markdown(answer)
        st.session_state.chat_history.append({"role": "assistant", "content": answer})
