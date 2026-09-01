# -*- coding: utf-8 -*-
"""
INEVOKE SARL — Dashboard de suivi des installations solaires
Version 2 : design modernisé, indicateurs de puissance enrichis,
rapports détaillés par installation (Excel / CSV) et analyses de performance.
"""

import os
import io
import base64
import numbers
from datetime import datetime

import numpy as np
import pandas as pd
import requests
import streamlit as st
import streamlit.components.v1 as components
import plotly.express as px
import plotly.graph_objects as go

try:
    from streamlit_autorefresh import st_autorefresh
except Exception:
    st_autorefresh = None

# ───────────────────────────────────────────────────────────────────────────
# Configuration de la page
# ───────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="INEVOKE — Suivi des installations solaires",
    page_icon="☀️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ───────────────────────────────────────────────────────────────────────────
# Identité visuelle INEVOKE
# ───────────────────────────────────────────────────────────────────────────
NAVY = "#0D47A1"
BLUE = "#2196F3"
SKY = "#42A5F5"
CYAN = "#00ACC1"
ORANGE = "#F9A825"
AMBER = "#FB8C00"
GREEN = "#2E7D32"
GREEN_L = "#43A047"
RED = "#C62828"
GREY = "#607D8B"
INK = "#12233F"
TXT = "#0A0A0A"   # texte des graphes : quasi-noir pour une lisibilité maximale
BG = "#F4F7FB"
CARD = "#FFFFFF"
LIGHT = "#E8F1FC"

# Palette catégorielle harmonisée
PALETTE = [BLUE, ORANGE, CYAN, GREEN_L, "#7E57C2", RED, "#26A69A", "#FFA726", "#5C6BC0"]
# Couleurs de statut cohérentes dans tout le tableau de bord
STATUT_COLORS = {
    "Opérationnel": GREEN_L,
    "En maintenance": ORANGE,
    "En panne": RED,
    "Hors service": GREY,
    "En cours d'installation": SKY,
    "Non renseigné": "#B0BEC5",
}

# ───────────────────────────────────────────────────────────────────────────
# Secrets Kobo
# ───────────────────────────────────────────────────────────────────────────
try:
    KOBO_API_URL = st.secrets.get("KOBO_API_URL", "https://kf.kobotoolbox.org")
    KOBO_ASSET_UID = st.secrets.get("KOBO_ASSET_UID", "")
    KOBO_API_TOKEN = st.secrets.get("KOBO_API_TOKEN", "")
    KOBO_FORM_LINK = st.secrets.get("KOBO_FORM_LINK", "")
except Exception:
    KOBO_API_URL = "https://kf.kobotoolbox.org"
    KOBO_ASSET_UID = ""
    KOBO_API_TOKEN = ""
    KOBO_FORM_LINK = ""

# Lien du formulaire de collecte Enketo (« Suivi des installations solaires déjà
# installées — INEVOKE SARL »). Modifiable dans les Secrets ; sinon, valeur par défaut.
DEFAULT_FORM_URL = "https://ee.kobotoolbox.org/x/8k1sXSOR"
KOBO_FORM_LINK = KOBO_FORM_LINK or DEFAULT_FORM_URL

# ───────────────────────────────────────────────────────────────────────────
# Feuille de style
# ───────────────────────────────────────────────────────────────────────────
st.markdown(
    f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Barlow:wght@300;400;500;600;700;800&family=Barlow+Semi+Condensed:wght@600;700&display=swap');

html, body, [class*="css"] {{ font-family:'Barlow', sans-serif; }}
.stApp {{ background:{BG}; }}
.block-container {{ padding-top:1.2rem; padding-bottom:2rem; max-width:1500px; }}

/* En-tête */
.main-header {{
    background:linear-gradient(120deg,{NAVY} 0%,{BLUE} 60%,{CYAN} 130%);
    padding:22px 30px; border-radius:20px; margin-bottom:1.3rem;
    display:flex; align-items:center; gap:22px;
    box-shadow:0 12px 34px rgba(13,71,161,.28);
}}
.main-header h1 {{ color:#fff; font-size:27px; font-weight:800; margin:0; letter-spacing:.2px; }}
.main-header p {{ color:rgba(255,255,255,.9); font-size:13.5px; margin:5px 0 0; }}
.header-badge {{
    margin-left:auto; text-align:right; color:#fff;
    background:rgba(255,255,255,.14); padding:10px 16px; border-radius:14px;
    font-size:12px; line-height:1.5; backdrop-filter:blur(4px);
}}
.header-badge b {{ font-size:15px; }}

/* Cartes KPI — hauteur fixe et zones réservées pour des cadres strictement identiques */
div[data-testid="stHorizontalBlock"] div[data-testid="column"] {{ display:flex; align-items:stretch; }}
.kpi-card {{
    background:{CARD}; border-radius:16px; padding:15px 16px;
    border-top:4px solid {BLUE}; box-shadow:0 4px 18px rgba(13,35,63,.07);
    height:148px; width:100%; box-sizing:border-box;
    position:relative; transition:transform .15s ease; overflow:hidden;
    display:flex; flex-direction:column; justify-content:flex-start;
}}
.kpi-card:hover {{ transform:translateY(-2px); }}
.kpi-card.orange {{ border-top-color:{ORANGE}; }}
.kpi-card.green  {{ border-top-color:{GREEN_L}; }}
.kpi-card.red    {{ border-top-color:{RED}; }}
.kpi-card.navy   {{ border-top-color:{NAVY}; }}
.kpi-card.cyan   {{ border-top-color:{CYAN}; }}
.kpi-label {{
    font-size:11px; color:#6B7A90; text-transform:uppercase; letter-spacing:.07em;
    font-weight:700; height:16px; line-height:16px; padding-right:26px;
    white-space:nowrap; overflow:hidden; text-overflow:ellipsis;
}}
.kpi-value {{
    font-size:25px; color:{INK}; font-weight:800; line-height:1.15;
    flex:1 1 auto; display:flex; align-items:center; margin-top:2px;
}}
.kpi-sub {{
    font-size:11.5px; color:#8894A6; font-weight:600; line-height:1.3;
    height:32px; display:flex; align-items:flex-end; overflow:hidden;
}}
.kpi-ico {{ position:absolute; top:13px; right:15px; font-size:18px; opacity:.32; }}

/* Titres de section */
.section-title {{
    font-size:16px; font-weight:800; color:{NAVY};
    border-left:5px solid {ORANGE}; padding-left:12px; margin:1.3rem 0 .8rem;
}}
.section-sub {{ font-size:12.5px; color:#7A889B; margin:-.4rem 0 .9rem 17px; }}

/* Bandeaux d'information */
.alert-box {{
    background:linear-gradient(90deg,{LIGHT},#F3F8FF);
    border-left:5px solid {BLUE}; border-radius:12px;
    padding:13px 18px; color:{NAVY}; font-size:13.5px; margin-bottom:1rem;
}}
.alert-warn {{ background:#FFF6E6; border-left-color:{ORANGE}; color:#8A5A00; }}

/* Fiche installation */
.install-card {{
    background:linear-gradient(135deg,#fff, #F7FBFF);
    border-radius:16px; padding:20px 22px;
    border-left:6px solid {BLUE}; box-shadow:0 4px 18px rgba(13,35,63,.08);
    margin-bottom:.4rem;
}}
.install-card h3 {{ color:{NAVY}; margin:0 0 4px; font-size:20px; }}
.chip {{
    display:inline-block; padding:3px 12px; border-radius:20px;
    font-size:12px; font-weight:700; color:#fff; margin-left:6px;
}}
.mini-metric {{
    background:{CARD}; border-radius:12px; padding:12px 14px; text-align:center;
    box-shadow:0 2px 10px rgba(13,35,63,.06); border-bottom:3px solid {LIGHT};
}}
.mini-metric .v {{ font-size:19px; font-weight:800; color:{INK}; }}
.mini-metric .l {{ font-size:10.5px; color:#7A889B; text-transform:uppercase; font-weight:700; letter-spacing:.04em; }}

/* Barre latérale */
section[data-testid="stSidebar"] {{ background:linear-gradient(180deg,{NAVY} 0%,#1565C0 100%); }}
section[data-testid="stSidebar"] * {{ color:#fff !important; }}
section[data-testid="stSidebar"] .stSelectbox label {{ font-size:12px; font-weight:700; }}

/* Boutons */
.stDownloadButton>button {{
    background:{ORANGE}; color:{INK}; border:none; border-radius:9px;
    font-weight:800; padding:.5rem 1rem;
}}
.stDownloadButton>button:hover {{ background:{AMBER}; color:#fff; }}
.stButton>button {{ background:{BLUE}; color:#fff; border:none; border-radius:9px; font-weight:800; }}
.stButton>button:hover {{ background:{NAVY}; }}

/* Onglets */
.stTabs [data-baseweb="tab-list"] {{ gap:4px; }}
.stTabs [data-baseweb="tab"] {{
    background:{CARD}; border-radius:10px 10px 0 0; padding:8px 16px;
    font-weight:700; font-size:13.5px; color:{GREY};
}}
.stTabs [aria-selected="true"] {{ background:{NAVY}; color:#fff !important; }}
[data-testid="stDataFrame"] {{ border-radius:12px; overflow:hidden; }}
</style>
""",
    unsafe_allow_html=True,
)


# ───────────────────────────────────────────────────────────────────────────
# Utilitaires
# ───────────────────────────────────────────────────────────────────────────
def image_to_base64(path):
    if not os.path.exists(path):
        return ""
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()


def fmt_num(v, decimals=1):
    try:
        if pd.isna(v):
            return "—"
        return f"{float(v):,.{decimals}f}".replace(",", " ")
    except Exception:
        return "—"


def empty_dataframe():
    cols = [
        "installation_id", "nom_site", "client", "type_site", "responsable_contact",
        "date_installation", "statut", "localisation_gps", "region", "ville", "commune",
        "quartier", "adresse", "photo_site", "nombre_panneaux", "puissance_unitaire_wc",
        "puissance_totale_kwc", "nombre_onduleurs", "puissance_onduleur_kva",
        "type_onduleur", "nombre_batteries", "capacite_batterie_kwh", "stockage_total_kwh",
        "tension_systeme_v", "puissance_max_kw", "production_estimee_kwh_jour",
        "production_reelle_kwh_jour", "production_mensuelle_kwh", "consommation_mensuelle_kwh",
        "autonomie_batterie_heures", "energie_injectee_kwh", "derniere_maintenance",
        "dernier_type_intervention", "dernier_probleme", "derniere_intervention",
        "dernier_technicien", "dernier_cout_intervention", "prochaine_maintenance",
        "observations", "marque_panneaux", "modele_panneaux", "marque_onduleur",
        "modele_onduleur", "marque_batteries", "modele_batteries", "numero_serie_principal",
        "agent_collecte", "date_collecte", "niveau_qualite", "commentaire_validation",
        "date_soumission", "latitude", "longitude",
    ]
    return pd.DataFrame(columns=cols)


def parse_geopoint(x):
    if isinstance(x, (list, tuple)) and len(x) >= 2:
        try:
            return float(x[0]), float(x[1])
        except Exception:
            return np.nan, np.nan
    if isinstance(x, str):
        parts = x.strip().replace(";", " ").split()
        if len(parts) >= 2:
            try:
                return float(parts[0]), float(parts[1])
            except Exception:
                pass
    return np.nan, np.nan


def safe_div(a, b):
    """Division vectorisée protégée contre les zéros et valeurs manquantes."""
    a = pd.to_numeric(a, errors="coerce")
    b = pd.to_numeric(b, errors="coerce")
    out = a / b.replace(0, np.nan)
    return out.replace([np.inf, -np.inf], np.nan)


# ───────────────────────────────────────────────────────────────────────────
# Normalisation + indicateurs dérivés
# ───────────────────────────────────────────────────────────────────────────
STATUT_LABELS = {
    "operationnel": "Opérationnel", "opérationnel": "Opérationnel",
    "maintenance": "En maintenance", "en maintenance": "En maintenance",
    "panne": "En panne", "en panne": "En panne",
    "hors_service": "Hors service", "hors service": "Hors service",
    "en_installation": "En cours d'installation",
}

# Décodage des codes de choix Kobo vers des libellés lisibles (feuille « choices » du XLSForm)
CHOICE_LABELS = {
    "type_site": {
        "residentiel": "Résidentiel", "commercial": "Commercial", "industriel": "Industriel",
        "agricole": "Agricole", "public": "Public / institutionnel", "autre": "Autre",
    },
    "region": {
        "abidjan": "Abidjan", "belier": "Bélier", "goh": "Gôh", "gbeke": "Gbêkê",
        "haut_sassandra": "Haut-Sassandra", "poro": "Poro", "san_pedro": "San-Pédro",
        "savanes": "Savanes", "yamoussoukro": "Yamoussoukro",
    },
    "ville": {
        "abidjan": "Abidjan", "yamoussoukro": "Yamoussoukro", "bouake": "Bouaké",
        "daloa": "Daloa", "gagnoa": "Gagnoa", "san_pedro": "San-Pédro",
        "korhogo": "Korhogo", "autres": "Autre ville",
    },
    "type_onduleur": {"hybride": "Hybride", "on_grid": "On-grid", "off_grid": "Off-grid"},
    "dernier_type_intervention": {
        "preventive": "Préventive", "corrective": "Corrective",
        "inspection": "Inspection / contrôle", "installation": "Installation / mise en service",
        "autre": "Autre",
    },
}


def decode_choice(series, table):
    """Traduit les codes de choix en libellés ; laisse la valeur telle quelle si inconnue."""
    lut = CHOICE_LABELS.get(table, {})
    return series.astype(str).str.strip().apply(
        lambda v: lut.get(v.lower(), v) if v and v.lower() != "nan" else "")


def enrich_indicators(df):
    """Calcule tous les indicateurs de puissance et de performance par installation."""
    # Recalcul de sécurité si Kobo n'a pas renvoyé les champs 'calculate'
    calc_power = df["nombre_panneaux"] * df["puissance_unitaire_wc"] / 1000
    df["puissance_totale_kwc"] = df["puissance_totale_kwc"].fillna(calc_power)
    calc_storage = df["nombre_batteries"] * df["capacite_batterie_kwh"]
    df["stockage_total_kwh"] = df["stockage_total_kwh"].fillna(calc_storage)

    # Ratio de dimensionnement onduleur / champ PV (kVA pour 1 kWc)
    df["ratio_dimensionnement"] = safe_div(df["puissance_onduleur_kva"], df["puissance_totale_kwc"])
    # Productible spécifique : énergie produite par kWc installé (kWh/kWc/jour)
    df["production_specifique"] = safe_div(df["production_reelle_kwh_jour"], df["puissance_totale_kwc"])
    # Ratio de performance : production réelle / production estimée (%)
    df["performance_ratio"] = safe_div(df["production_reelle_kwh_jour"], df["production_estimee_kwh_jour"]) * 100
    # Taux de couverture des besoins : production / consommation mensuelle (%)
    df["taux_couverture"] = safe_div(df["production_mensuelle_kwh"], df["consommation_mensuelle_kwh"]) * 100
    # Ratio de stockage : kWh de batterie pour 1 kWc PV
    df["ratio_stockage"] = safe_div(df["stockage_total_kwh"], df["puissance_totale_kwc"])
    # Production annuelle estimée à partir de la production mensuelle (ou journalière)
    prod_annuelle = df["production_mensuelle_kwh"] * 12
    prod_annuelle = prod_annuelle.fillna(df["production_estimee_kwh_jour"] * 365)
    df["production_annuelle_kwh"] = prod_annuelle
    return df


def normalize_columns(df):
    if df is None or len(df) == 0:
        return enrich_indicators(empty_dataframe())

    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    # Votre XLSForm utilise des groupes (begin_group). L'API Kobo renvoie donc les
    # champs préfixés par le chemin du groupe, ex. « identification/nom_site » ou
    # « photovoltaique/nombre_panneaux ». On ne conserve que le nom final du champ.
    # Les champs méta de Kobo (commençant par « _ ») sont laissés intacts.
    def _strip_group(c):
        return c if c.startswith("_") else c.rsplit("/", 1)[-1]

    df = df.rename(columns={c: _strip_group(c) for c in df.columns})
    # Supprime d'éventuels doublons de colonnes nés du retrait des préfixes
    df = df.loc[:, ~df.columns.duplicated()]

    # Alias des champs techniques Kobo, sans jamais écraser un champ déjà présent
    if "installation_id" not in df.columns and "_id" in df.columns:
        df["installation_id"] = df["_id"]
    if "date_soumission" not in df.columns and "_submission_time" in df.columns:
        df["date_soumission"] = df["_submission_time"]
    if "geolocation" not in df.columns and "_geolocation" in df.columns:
        df["geolocation"] = df["_geolocation"]

    for col in empty_dataframe().columns:
        if col not in df.columns:
            df[col] = np.nan

    # GPS — on préserve toute latitude/longitude déjà fournie et on ne dérive les
    # coordonnées d'un champ geopoint ('geolocation' ou 'localisation_gps') que
    # lorsqu'elles sont absentes, afin de ne jamais écraser une position valide.
    lat_existante = pd.to_numeric(df.get("latitude"), errors="coerce")
    lon_existante = pd.to_numeric(df.get("longitude"), errors="coerce")
    a_completer = lat_existante.isna() | lon_existante.isna()

    def _col_a_du_contenu(col):
        if col not in df.columns:
            return False
        s = df[col].astype(str).str.strip().replace({"nan": "", "None": "", "NaT": ""})
        return s.ne("").any()

    source_geo = "geolocation" if _col_a_du_contenu("geolocation") else \
                 "localisation_gps" if _col_a_du_contenu("localisation_gps") else None

    if source_geo:
        coords = df[source_geo].apply(parse_geopoint)
        plat = coords.apply(lambda t: t[0])
        plon = coords.apply(lambda t: t[1])
        df["latitude"] = np.where(a_completer, plat, lat_existante)
        df["longitude"] = np.where(a_completer, plon, lon_existante)
    else:
        df["latitude"] = lat_existante
        df["longitude"] = lon_existante

    date_cols = ["date_installation", "derniere_maintenance", "prochaine_maintenance",
                 "date_collecte", "date_soumission"]
    num_cols = [
        "nombre_panneaux", "puissance_unitaire_wc", "puissance_totale_kwc",
        "nombre_onduleurs", "puissance_onduleur_kva", "nombre_batteries",
        "capacite_batterie_kwh", "stockage_total_kwh", "tension_systeme_v", "puissance_max_kw",
        "production_estimee_kwh_jour", "production_reelle_kwh_jour", "production_mensuelle_kwh",
        "consommation_mensuelle_kwh", "autonomie_batterie_heures", "energie_injectee_kwh",
        "dernier_cout_intervention",
    ]
    for col in date_cols:
        df[col] = pd.to_datetime(df[col], errors="coerce")
    for col in num_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df = enrich_indicators(df)

    text_cols = [c for c in empty_dataframe().columns
                 if c not in date_cols + num_cols + ["latitude", "longitude", "localisation_gps"]]
    for col in text_cols:
        df[col] = df[col].fillna("").astype(str).str.strip()

    # Décodage des codes de choix Kobo en libellés lisibles
    df["statut"] = df["statut"].apply(
        lambda s: STATUT_LABELS.get(str(s).strip().lower(), s) if str(s).strip() else "Non renseigné")
    df["region"] = decode_choice(df["region"], "region").replace("", "Non renseignée")
    df["ville"] = decode_choice(df["ville"], "ville").replace("", "Non renseignée")
    df["type_site"] = decode_choice(df["type_site"], "type_site").replace("", "Non renseigné")
    df["type_onduleur"] = decode_choice(df["type_onduleur"], "type_onduleur")
    df["dernier_type_intervention"] = decode_choice(
        df["dernier_type_intervention"], "dernier_type_intervention")

    df["annee"] = df["date_installation"].dt.year
    df["mois"] = df["date_installation"].dt.to_period("M").astype(str)
    return df


# ───────────────────────────────────────────────────────────────────────────
# Connexion Kobo
# ───────────────────────────────────────────────────────────────────────────
@st.cache_data(ttl=60, show_spinner=False)
def fetch_kobo_data(api_url, asset_uid, token):
    if not asset_uid or not token:
        return enrich_indicators(empty_dataframe())
    url = f"{api_url.rstrip('/')}/api/v2/assets/{asset_uid}/data.json"
    headers = {"Authorization": f"Token {token}"}
    rows, next_url = [], url
    try:
        while next_url:
            response = requests.get(next_url, headers=headers, timeout=35)
            response.raise_for_status()
            payload = response.json()
            rows.extend(payload.get("results", []))
            next_url = payload.get("next")
            if next_url and next_url.startswith("/"):
                next_url = api_url.rstrip("/") + next_url
        return normalize_columns(pd.DataFrame(rows))
    except Exception as e:
        st.error(f"Erreur de connexion à Kobo : {e}")
        return enrich_indicators(empty_dataframe())


# ───────────────────────────────────────────────────────────────────────────
# Jeu de démonstration (enrichi avec production / consommation)
# ───────────────────────────────────────────────────────────────────────────
def demo_dataframe():
    data = [
        # id, site, client, type, region, ville, lat, lon, np, wc, ond, kva, ond_type, nb_bat, cap_bat, statut, commune, prod_est_j, prod_reel_j, prod_mois, conso_mois, inject, autonomie
        ["INV-SOL-0001", "Usine Yopougon", "AgroPlus SA", "industriel", "Abidjan", "Abidjan", 5.345, -4.083, 120, 550, 5, 90, "on_grid", 0, 0, "operationnel", "Yopougon", 320, 305, 9150, 14200, 1800, 0],
        ["INV-SOL-0002", "Dépôt Gagnoa", "AgroPlus SA", "industriel", "Gôh", "Gagnoa", 6.131, -5.951, 96, 550, 4, 66, "on_grid", 0, 0, "operationnel", "Gagnoa", 255, 240, 7200, 9000, 900, 0],
        ["INV-SOL-0003", "Hôtel Yamoussoukro", "Ivoire Hôtels", "commercial", "Yamoussoukro", "Yamoussoukro", 6.816, -5.276, 80, 550, 4, 50, "hybride", 12, 5.12, "operationnel", "Centre", 215, 198, 5940, 7800, 350, 14],
        ["INV-SOL-0004", "Hôtel Cocody", "Ivoire Hôtels", "commercial", "Abidjan", "Abidjan", 5.359, -3.998, 24, 550, 2, 15, "hybride", 4, 5.12, "maintenance", "Cocody", 62, 48, 1440, 2100, 120, 8],
        ["INV-SOL-0005", "Clinique San-Pédro", "Ministère Santé", "public", "San-Pédro", "San-Pédro", 4.748, -6.637, 30, 550, 2, 16, "hybride", 8, 5.12, "operationnel", "San-Pédro", 82, 79, 2370, 2200, 40, 16],
        ["INV-SOL-0006", "Centre Bouaké", "Ministère Santé", "public", "Gbêkê", "Bouaké", 7.690, -5.030, 60, 550, 3, 33, "hybride", 8, 5.12, "panne", "Bouaké", 165, 0, 0, 4100, 0, 12],
        ["INV-SOL-0007", "Ferme Daloa", "Coopérative Café", "agricole", "Haut-Sassandra", "Daloa", 6.880, -6.450, 36, 450, 2, 18, "off_grid", 10, 10.0, "operationnel", "Daloa", 92, 88, 2640, 2500, 0, 20],
        ["INV-SOL-0008", "Marché Korhogo", "Mairie Korhogo", "public", "Poro", "Korhogo", 9.458, -5.629, 48, 550, 3, 30, "on_grid", 6, 5.12, "operationnel", "Korhogo", 130, 121, 3630, 3200, 300, 6],
        ["INV-SOL-0009", "Villa Riviera", "Groupe Koffi", "residentiel", "Abidjan", "Abidjan", 5.360, -3.980, 12, 450, 1, 8, "hybride", 4, 5.12, "operationnel", "Cocody", 30, 27, 810, 900, 0, 10],
        ["INV-SOL-0010", "École Abobo", "Éducation Nationale", "public", "Abidjan", "Abidjan", 5.416, -4.016, 20, 450, 1, 10, "off_grid", 4, 5.12, "operationnel", "Abobo", 48, 44, 1320, 1200, 0, 12],
    ]
    cols = ["installation_id", "nom_site", "client", "type_site", "region", "ville", "latitude", "longitude",
            "nombre_panneaux", "puissance_unitaire_wc", "nombre_onduleurs", "puissance_onduleur_kva",
            "type_onduleur", "nombre_batteries", "capacite_batterie_kwh", "statut", "commune",
            "production_estimee_kwh_jour", "production_reelle_kwh_jour", "production_mensuelle_kwh",
            "consommation_mensuelle_kwh", "energie_injectee_kwh", "autonomie_batterie_heures"]
    d = pd.DataFrame(data, columns=cols)
    d["puissance_totale_kwc"] = d["nombre_panneaux"] * d["puissance_unitaire_wc"] / 1000
    d["stockage_total_kwh"] = d["nombre_batteries"] * d["capacite_batterie_kwh"]
    d["tension_systeme_v"] = 48
    d["puissance_max_kw"] = d["puissance_onduleur_kva"] * 0.9
    dates = pd.to_datetime([
        "2024-02-10", "2024-05-22", "2024-08-15", "2024-11-03", "2025-01-18",
        "2025-03-27", "2025-06-09", "2025-08-30", "2025-11-14", "2026-01-20"])
    d["date_installation"] = dates
    d["prochaine_maintenance"] = d["date_installation"] + pd.Timedelta(days=280)
    d["derniere_maintenance"] = d["date_installation"] + pd.Timedelta(days=90)
    d["dernier_technicien"] = "Équipe INEVOKE"
    for c in empty_dataframe().columns:
        if c not in d.columns:
            d[c] = np.nan
    return normalize_columns(d)


# ───────────────────────────────────────────────────────────────────────────
# Génération du rapport Excel multi-feuilles
# ───────────────────────────────────────────────────────────────────────────
def build_excel_report(df):
    """Construit un rapport Excel des puissances par installation."""
    export = df.copy()
    colonnes = {
        "installation_id": "ID installation",
        "nom_site": "Site",
        "client": "Client",
        "type_site": "Type de site",
        "ville": "Ville",
        "region": "Région",
        "statut": "Statut",
        "nombre_panneaux": "Nb panneaux",
        "puissance_unitaire_wc": "Puissance unitaire (Wc)",
        "puissance_totale_kwc": "Puissance PV (kWc)",
        "nombre_onduleurs": "Nb onduleurs",
        "puissance_onduleur_kva": "Puissance onduleurs (kVA)",
        "type_onduleur": "Type onduleur",
        "ratio_dimensionnement": "Ratio kVA/kWc",
        "nombre_batteries": "Nb batteries",
        "stockage_total_kwh": "Stockage (kWh)",
        "ratio_stockage": "Stockage/kWc",
        "production_estimee_kwh_jour": "Prod. estimée (kWh/j)",
        "production_reelle_kwh_jour": "Prod. réelle (kWh/j)",
        "production_specifique": "Productible (kWh/kWc/j)",
        "performance_ratio": "Ratio perf. (%)",
        "production_mensuelle_kwh": "Prod. mensuelle (kWh)",
        "consommation_mensuelle_kwh": "Conso mensuelle (kWh)",
        "taux_couverture": "Couverture besoins (%)",
        "production_annuelle_kwh": "Prod. annuelle est. (kWh)",
        "energie_injectee_kwh": "Énergie injectée (kWh)",
    }
    detail = export[[c for c in colonnes if c in export.columns]].rename(columns=colonnes)

    # Feuille de synthèse
    synth = pd.DataFrame({
        "Indicateur": [
            "Nombre d'installations", "Puissance PV totale (kWc)", "Puissance onduleurs totale (kVA)",
            "Capacité de stockage totale (kWh)", "Nombre total de panneaux",
            "Production réelle cumulée (kWh/j)", "Production annuelle estimée (kWh)",
            "Ratio de performance moyen (%)", "Productible spécifique moyen (kWh/kWc/j)",
            "Installations opérationnelles", "Installations en maintenance", "Installations en panne",
        ],
        "Valeur": [
            len(export),
            round(export["puissance_totale_kwc"].sum(), 2),
            round(export["puissance_onduleur_kva"].sum(), 2),
            round(export["stockage_total_kwh"].sum(), 2),
            int(export["nombre_panneaux"].sum()),
            round(export["production_reelle_kwh_jour"].sum(), 1),
            round(export["production_annuelle_kwh"].sum(), 0),
            round(export["performance_ratio"].mean(), 1),
            round(export["production_specifique"].mean(), 2),
            int(export["statut"].eq("Opérationnel").sum()),
            int(export["statut"].eq("En maintenance").sum()),
            int(export["statut"].eq("En panne").sum()),
        ],
    })

    # Synthèse par ville
    par_ville = export.groupby("ville").agg(
        Installations=("installation_id", "count"),
        Puissance_PV_kWc=("puissance_totale_kwc", "sum"),
        Onduleurs_kVA=("puissance_onduleur_kva", "sum"),
        Stockage_kWh=("stockage_total_kwh", "sum"),
        Prod_reelle_kWh_j=("production_reelle_kwh_jour", "sum"),
    ).round(2).reset_index()

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        synth.to_excel(writer, sheet_name="Synthèse", index=False)
        par_ville.to_excel(writer, sheet_name="Par ville", index=False)
        detail.to_excel(writer, sheet_name="Détail par installation", index=False)
    buffer.seek(0)
    return buffer.getvalue()


def kpi_card(label, value, color_class="blue", sub="", icon=""):
    # La zone sous-titre est toujours rendue (même vide) afin que toutes les
    # cartes conservent exactement la même structure et la même hauteur.
    ico = f'<div class="kpi-ico">{icon}</div>' if icon else ""
    st.markdown(
        f'<div class="kpi-card {color_class}">{ico}'
        f'<div class="kpi-label">{label}</div>'
        f'<div class="kpi-value">{value}</div>'
        f'<div class="kpi-sub">{sub}</div></div>',
        unsafe_allow_html=True,
    )


def styled_fig(fig, height=380, legend=True):
    fig.update_layout(
        height=height, plot_bgcolor="white", paper_bgcolor="white",
        font=dict(family="Barlow, sans-serif", size=14, color=TXT),
        title_font=dict(family="Barlow, sans-serif", size=17, color=TXT),
        margin=dict(l=10, r=10, t=34, b=10),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0,
                    font=dict(color=TXT, size=13)) if legend else dict(),
    )
    # Évite l'affichage d'un titre « indéfini » quand aucun titre n'est défini
    if fig.layout.title.text is None:
        fig.update_layout(title_text="")

    axis_common = dict(
        title_font=dict(color=TXT, size=14),
        tickfont=dict(color=TXT, size=13),
        linecolor="#5A6B82", linewidth=1.4,
    )
    fig.update_xaxes(showgrid=False, **axis_common)
    fig.update_yaxes(showgrid=True, gridcolor="#E6EBF2", zeroline=False, **axis_common)
    # Valeurs affichées sur les barres : noires, nettes et jamais rognées
    fig.update_traces(selector=dict(type="bar"),
                      textfont=dict(color=TXT, size=13), cliponaxis=False)

    # Étend l'axe des barres horizontales pour que les étiquettes de bout de barre
    # (ex. « 118,8 kWc ») restent entièrement lisibles et non coupées au bord droit.
    xmax = 0.0
    for tr in fig.data:
        if getattr(tr, "type", "") == "bar" and getattr(tr, "orientation", None) == "h":
            xs_raw = tr.x if tr.x is not None else []
            xs = [float(v) for v in xs_raw if isinstance(v, numbers.Number)]
            if xs:
                xmax = max(xmax, max(xs))
    if xmax > 0:
        fig.update_xaxes(range=[0, xmax * 1.18])
    return fig


# ───────────────────────────────────────────────────────────────────────────
# En-tête
# ───────────────────────────────────────────────────────────────────────────
logo_b64 = image_to_base64("assets/logo_inevoke.jpeg")
logo_html = (f'<img src="data:image/jpeg;base64,{logo_b64}" '
             f'style="height:66px;background:#fff;border-radius:12px;padding:6px 10px;">'
             if logo_b64 else "☀️")

st.markdown(
    f"""
<div class="main-header">
    <div>{logo_html}</div>
    <div>
        <h1>Suivi des installations solaires</h1>
        <p>INEVOKE SARL · Centralisation, cartographie et analyse des puissances installées</p>
    </div>
    <div class="header-badge">
        <b>{datetime.now().strftime("%d/%m/%Y")}</b><br>{datetime.now().strftime("%H:%M")}
    </div>
</div>
""",
    unsafe_allow_html=True,
)

# ───────────────────────────────────────────────────────────────────────────
# Barre latérale
# ───────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🔄 Actualisation")
    refresh_seconds = st.selectbox("Fréquence", [30, 60, 120, 300], index=1,
                                   format_func=lambda x: f"Toutes les {x} s")
    if st_autorefresh is not None:
        st_autorefresh(interval=refresh_seconds * 1000, key="auto_refresh_installations")

    st.markdown("### 🔗 Formulaire Kobo")
    if KOBO_FORM_LINK:
        st.markdown(f"[Ouvrir le formulaire terrain]({KOBO_FORM_LINK})")
    else:
        st.caption("Lien Kobo à renseigner dans les Secrets.")
    if st.button("Forcer l'actualisation"):
        st.cache_data.clear()
        st.rerun()

    st.markdown("### 🗂️ Source des données")
    st.caption("Utilisé uniquement tant que Kobo n'est pas connecté.")
    source_hors_kobo = st.radio(
        "Affichage sans connexion Kobo",
        ["Tableau vide (démarrer à zéro)", "Données de démonstration"],
        index=0, label_visibility="collapsed")


# ═══════════════════════════════════════════════════════════════════════════
#  INTERFACE — Suivi des puissances DÉJÀ INSTALLÉES, par CLIENT
# ═══════════════════════════════════════════════════════════════════════════
df_all = fetch_kobo_data(KOBO_API_URL, KOBO_ASSET_UID, KOBO_API_TOKEN)
kobo_connecte = len(df_all) > 0

# Hors connexion Kobo, l'utilisateur choisit entre un tableau vierge (démarrage à
# zéro) et les données de démonstration. Par défaut : tableau vierge.
if kobo_connecte:
    mode = "kobo"
elif source_hors_kobo.startswith("Données de démonstration"):
    mode = "demo"
    df_all = demo_dataframe()
else:
    mode = "vide"
    df_all = enrich_indicators(empty_dataframe())

# On ne conserve que les installations RÉELLEMENT installées (on écarte
# celles encore « en cours d'installation »), conformément à l'objectif :
# suivi des puissances déjà installées pour chaque client.
non_installe = df_all["statut"].eq("En cours d'installation").sum()
df_pose = df_all[df_all["statut"] != "En cours d'installation"].copy()

# Client non renseigné → libellé lisible
df_pose["client"] = df_pose["client"].replace("", "Client non renseigné")

with st.sidebar:
    st.markdown("---")
    st.markdown("### Filtres")
    df = df_pose.copy()

    def _filter(col, label, all_label="Tous"):
        global df
        opts = [all_label] + sorted([v for v in df_pose[col].dropna().astype(str).unique() if v])
        sel = st.selectbox(label, opts, key=f"f_{col}")
        if sel != all_label:
            df = df[df[col].astype(str) == sel]

    _filter("client", "Client", "Tous")           # filtre principal
    _filter("ville", "Ville", "Toutes")
    _filter("statut", "Statut")
    _filter("type_site", "Type de site")
    _filter("type_onduleur", "Type d'onduleur")

    st.markdown("---")
    st.caption(f"**{len(df)}** installation(s) · **{df['client'].nunique()}** client(s)")

# Bandeau de mode
if mode == "demo":
    st.markdown(
        '<div class="alert-box alert-warn"><b> Mode démonstration.</b> Données fictives '
        'illustrant le suivi des puissances installées par client. Passez sur « Tableau vide » '
        'dans la barre latérale, ou renseignez les Secrets Kobo pour vos données réelles.</div>',
        unsafe_allow_html=True)
elif mode == "vide":
    st.markdown(
        '<div class="alert-box"><b> Tableau vierge.</b> Aucune donnée chargée : les compteurs '
        'sont à zéro. Saisissez vos installations depuis l\'onglet « 🛰️ Collecte & saisie », '
        'puis cliquez sur « Forcer l\'actualisation ». (Connectez les Secrets Kobo pour un '
        'affichage automatique des soumissions.)</div>', unsafe_allow_html=True)
else:
    st.markdown(
        f'<div class="alert-box"><b>🛰️ Connecté à KoboToolbox.</b> {len(df_pose)} installation(s) '
        f'déjà posée(s)' + (f' · {non_installe} en cours d\'installation exclue(s)' if non_installe else '')
        + f' · synchro {datetime.now():%d/%m/%Y %H:%M}.</div>', unsafe_allow_html=True)

# ───────────────────────────────────────────────────────────────────────────
# KPI — orientés client / puissance installée
# ───────────────────────────────────────────────────────────────────────────
n_clients = df["client"].nunique()
pv_kwc = df["puissance_totale_kwc"].sum()
inv_kva = df["puissance_onduleur_kva"].sum()
storage_kwh = df["stockage_total_kwh"].sum()
prod_reel = df["production_reelle_kwh_jour"].sum()
operational = df["statut"].eq("Opérationnel").sum()
n = max(len(df), 1)
dispo = operational / n * 100
pv_par_client = pv_kwc / max(n_clients, 1)

k = st.columns(8)
with k[0]: kpi_card("Clients", n_clients, "navy", "portefeuille suivi", "")
with k[1]: kpi_card("Installations", len(df), "blue", f"{df['ville'].nunique()} ville(s)", "")
with k[2]: kpi_card("Puissance installée", f"{fmt_num(pv_kwc, 1)} kWc", "orange", f"{fmt_num(pv_kwc/1000, 2)} MWc", "")
with k[3]: kpi_card("Moyenne / client", f"{fmt_num(pv_par_client, 1)} kWc", "cyan", "puissance PV", "")
with k[4]: kpi_card("Onduleurs", f"{fmt_num(inv_kva, 1)} kVA", "cyan", f"{fmt_num(inv_kva/1000, 2)} MVA", "")
with k[5]: kpi_card("Stockage", f"{fmt_num(storage_kwh, 1)} kWh", "green", f"{fmt_num(storage_kwh/1000, 2)} MWh", "")
with k[6]: kpi_card("Prod. réelle", f"{fmt_num(prod_reel, 0)} kWh/j", "orange", "cumul terrain", "")
with k[7]: kpi_card("Disponibilité", f"{fmt_num(dispo, 0)} %", "green" if dispo >= 80 else "orange", f"{operational} opérationnelle(s)", "")


# ───────────────────────────────────────────────────────────────────────────
# Agrégat par client (réutilisé par plusieurs onglets)
# ───────────────────────────────────────────────────────────────────────────
def agg_clients(d):
    a = d.groupby("client").agg(
        Installations=("installation_id", "count"),
        PV_kWc=("puissance_totale_kwc", "sum"),
        Onduleurs_kVA=("puissance_onduleur_kva", "sum"),
        Stockage_kWh=("stockage_total_kwh", "sum"),
        Prod_reelle=("production_reelle_kwh_jour", "sum"),
        Operationnelles=("statut", lambda s: (s == "Opérationnel").sum()),
    ).reset_index()
    tot = a["PV_kWc"].sum() or 1
    a["Part_%"] = a["PV_kWc"] / tot * 100
    a["Dispo_%"] = a["Operationnelles"] / a["Installations"] * 100
    return a.sort_values("PV_kWc", ascending=False)


# ───────────────────────────────────────────────────────────────────────────
# Onglets
# ───────────────────────────────────────────────────────────────────────────
tabs = st.tabs([
    " Par client",
    " Puissances & équipements",
    " Performance énergétique",
    " Carte",
    " Rapport par installation",
    " Fiche client",
    " Données brutes",
    " Collecte & saisie",
])

# ============================ 0. PAR CLIENT ============================
with tabs[0]:
    if df.empty:
        st.markdown(
            f"""
            <div style="background:#fff;border:2px dashed {BLUE};border-radius:16px;
                        padding:34px;text-align:center;color:{NAVY};margin-top:.5rem;">
              <div style="font-size:34px;">📭</div>
              <h3 style="margin:.3rem 0;color:{NAVY};">Aucune installation enregistrée</h3>
              <p style="color:#6B7A90;max-width:560px;margin:.4rem auto;">
                 Le tableau de bord est à zéro. Rendez-vous dans l'onglet
                 <b>« 🛰️ Collecte &amp; saisie »</b> pour enregistrer vos installations,
                 puis cliquez sur <b>« Forcer l'actualisation »</b> dans la barre latérale.<br>
                 Pour tester l'affichage, choisissez <b>« Données de démonstration »</b>
                 dans la barre latérale.
              </p>
            </div>
            """, unsafe_allow_html=True)
    else:
        st.markdown("<div class='section-title'>Puissance installée par client</div>", unsafe_allow_html=True)
        cli = agg_clients(df)
        c1, c2 = st.columns([1.45, 1])
        with c1:
            cc = cli.sort_values("PV_kWc")
            fig = px.bar(cc, x="PV_kWc", y="client", orientation="h", text="PV_kWc",
                         color="PV_kWc", color_continuous_scale=["#BBDEFB", BLUE, NAVY],
                         hover_data={"Installations": True, "Onduleurs_kVA": ":.1f",
                                     "Stockage_kWh": ":.1f", "PV_kWc": ":.1f"})
            fig.update_traces(texttemplate="%{text:.1f} kWc", textposition="outside")
            fig.update_layout(coloraxis_showscale=False, xaxis_title="Puissance PV installée (kWc)", yaxis_title="")
            st.plotly_chart(styled_fig(fig, 430, legend=False), use_container_width=True)
        with c2:
            fig = px.pie(cli, values="PV_kWc", names="client", hole=.55,
                         color_discrete_sequence=PALETTE)
            fig.update_traces(textposition="outside", textinfo="label+percent",
                              textfont=dict(color=TXT, size=13), automargin=True,
                              marker=dict(line=dict(color="white", width=2)))
            fig.update_layout(title="Répartition de la puissance", showlegend=False)
            st.plotly_chart(styled_fig(fig, 430, legend=False), use_container_width=True)

        st.markdown("<div class='section-title'>Synthèse par client</div>", unsafe_allow_html=True)
        show = cli.rename(columns={
            "client": "Client", "Installations": "Sites", "PV_kWc": "PV (kWc)",
            "Onduleurs_kVA": "Onduleurs (kVA)", "Stockage_kWh": "Stockage (kWh)",
            "Prod_reelle": "Prod. réelle (kWh/j)", "Part_%": "Part du parc (%)", "Dispo_%": "Dispo. (%)"})
        show = show[["Client", "Sites", "PV (kWc)", "Part du parc (%)", "Onduleurs (kVA)",
                     "Stockage (kWh)", "Prod. réelle (kWh/j)", "Dispo. (%)"]]
        st.dataframe(
            show.style.format({
                "PV (kWc)": "{:.1f}", "Part du parc (%)": "{:.1f}", "Onduleurs (kVA)": "{:.1f}",
                "Stockage (kWh)": "{:.1f}", "Prod. réelle (kWh/j)": "{:.0f}", "Dispo. (%)": "{:.0f}",
            }, na_rep="—").background_gradient(subset=["PV (kWc)"], cmap="Blues"),
            use_container_width=True, hide_index=True)
        st.download_button("📥 Télécharger la synthèse clients (CSV)",
                           show.to_csv(index=False).encode("utf-8-sig"),
                           f"synthese_clients_inevoke_{datetime.now():%Y%m%d}.csv", "text/csv")

        st.markdown("<div class='section-title'>Détail client → sites (poids en puissance)</div>", unsafe_allow_html=True)
        tm = df[df["puissance_totale_kwc"] > 0].copy()
        if len(tm):
            fig = px.treemap(tm, path=[px.Constant("Parc INEVOKE"), "client", "nom_site"],
                             values="puissance_totale_kwc", color="puissance_totale_kwc",
                             color_continuous_scale=["#E3F2FD", "#90CAF9", "#42A5F5"])
            fig.update_traces(texttemplate="%{label}<br>%{value:.1f} kWc",
                              textfont=dict(color=TXT, size=14),
                              insidetextfont=dict(color=TXT, size=14),
                              marker=dict(line=dict(color="white", width=2)))
            fig.update_layout(height=430, margin=dict(l=6, r=6, t=6, b=6),
                              font=dict(color=TXT), coloraxis_showscale=False,
                              paper_bgcolor="white", plot_bgcolor="white")
            st.plotly_chart(fig, use_container_width=True)

# ==================== 1. PUISSANCES & ÉQUIPEMENTS ====================
with tabs[1]:
    st.markdown("<div class='section-title'>PV vs onduleurs par client</div>", unsafe_allow_html=True)
    cli = agg_clients(df)
    c1, c2 = st.columns(2)
    with c1:
        fig = go.Figure()
        fig.add_trace(go.Bar(x=cli["client"], y=cli["PV_kWc"], name="PV (kWc)", marker_color=BLUE))
        fig.add_trace(go.Bar(x=cli["client"], y=cli["Onduleurs_kVA"], name="Onduleurs (kVA)", marker_color=ORANGE))
        fig.update_layout(barmode="group", yaxis_title="Puissance", xaxis_title="")
        st.plotly_chart(styled_fig(fig, 400), use_container_width=True)
    with c2:
        cc = cli.sort_values("Stockage_kWh")
        fig = px.bar(cc, x="Stockage_kWh", y="client", orientation="h", text="Stockage_kWh",
                     color_discrete_sequence=[GREEN_L])
        fig.update_traces(texttemplate="%{text:.1f} kWh", textposition="outside")
        fig.update_layout(xaxis_title="Stockage (kWh)", yaxis_title="", title="Capacité de stockage")
        st.plotly_chart(styled_fig(fig, 400, legend=False), use_container_width=True)

    st.markdown("<div class='section-title'>Dimensionnement onduleur / champ PV</div>", unsafe_allow_html=True)
    st.markdown("<div class='section-sub'>Ratio kVA/kWc proche de 0,8–1,1 = équilibré ; "
                "chaque point est une installation, coloré par client.</div>", unsafe_allow_html=True)
    scat = df[df["puissance_totale_kwc"] > 0].copy()
    if len(scat):
        fig = px.scatter(scat, x="puissance_totale_kwc", y="puissance_onduleur_kva",
                         size="nombre_panneaux", color="client", color_discrete_sequence=PALETTE,
                         hover_name="nom_site", size_max=30,
                         labels={"puissance_totale_kwc": "Puissance PV (kWc)",
                                 "puissance_onduleur_kva": "Onduleurs (kVA)"})
        maxv = max(scat["puissance_totale_kwc"].max(), scat["puissance_onduleur_kva"].max())
        fig.add_trace(go.Scatter(x=[0, maxv], y=[0, maxv], mode="lines",
                                 line=dict(color=GREY, dash="dash"), name="Ratio 1:1"))
        st.plotly_chart(styled_fig(fig, 420), use_container_width=True)

# ==================== 2. PERFORMANCE ====================
with tabs[2]:
    st.markdown("<div class='section-title'>Production estimée vs réelle par client</div>", unsafe_allow_html=True)
    perf = df.groupby("client", as_index=False).agg(
        est=("production_estimee_kwh_jour", "sum"),
        reel=("production_reelle_kwh_jour", "sum"))
    perf = perf[(perf["est"] > 0) | (perf["reel"] > 0)]
    if len(perf):
        c1, c2 = st.columns([1.4, 1])
        with c1:
            fig = go.Figure()
            fig.add_trace(go.Bar(x=perf["client"], y=perf["est"], name="Estimée (kWh/j)", marker_color=SKY))
            fig.add_trace(go.Bar(x=perf["client"], y=perf["reel"], name="Réelle (kWh/j)", marker_color=ORANGE))
            fig.update_layout(barmode="group", yaxis_title="kWh/jour", xaxis_title="")
            st.plotly_chart(styled_fig(fig, 400), use_container_width=True)
        with c2:
            pr = df.dropna(subset=["performance_ratio"])
            avg = pr["performance_ratio"].mean() if len(pr) else 0
            fig = go.Figure(go.Indicator(
                mode="gauge+number", value=avg,
                number={"suffix": " %", "font": {"color": TXT, "size": 44}},
                title={"text": "Ratio de performance moyen", "font": {"color": TXT, "size": 16}},
                gauge={"axis": {"range": [0, 120], "tickfont": {"color": TXT, "size": 13}},
                       "bar": {"color": NAVY},
                       "steps": [{"range": [0, 70], "color": "#FFCDD2"},
                                 {"range": [70, 90], "color": "#FFF3CD"},
                                 {"range": [90, 120], "color": "#C8E6C9"}]}))
            fig.update_layout(height=400, margin=dict(l=55, r=60, t=40, b=10),
                              font=dict(color=TXT), paper_bgcolor="white", plot_bgcolor="white")
            st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Renseignez la production estimée et réelle dans Kobo pour activer ces analyses.")

    st.markdown("<div class='section-title'>Productible spécifique par installation (kWh/kWc/jour)</div>", unsafe_allow_html=True)
    ps = df.dropna(subset=["production_specifique"]).sort_values("production_specifique")
    if len(ps):
        fig = px.bar(ps, x="production_specifique", y="nom_site", orientation="h",
                     text="production_specifique", color="client", color_discrete_sequence=PALETTE)
        fig.update_traces(texttemplate="%{text:.2f}", textposition="outside")
        fig.update_layout(xaxis_title="kWh/kWc/jour", yaxis_title="")
        st.plotly_chart(styled_fig(fig, 420), use_container_width=True)

# ==================== 3. CARTE ====================
with tabs[3]:
    st.markdown("<div class='section-title'>Cartographie des installations</div>", unsafe_allow_html=True)
    geo = df.copy()
    geo["latitude"] = pd.to_numeric(geo["latitude"], errors="coerce")
    geo["longitude"] = pd.to_numeric(geo["longitude"], errors="coerce")
    geo = geo.dropna(subset=["latitude", "longitude"])
    geo = geo[geo["latitude"].between(-90, 90) & geo["longitude"].between(-180, 180)
              & ~((geo["latitude"].abs() < 0.01) & (geo["longitude"].abs() < 0.01))].copy()
    sites_sans_gps = len(df) - len(geo)

    if len(geo):
        geo["taille"] = pd.to_numeric(geo["puissance_totale_kwc"], errors="coerce").fillna(1).clip(lower=1)
        lat_c, lon_c = float(geo["latitude"].mean()), float(geo["longitude"].mean())
        span = max(geo["latitude"].max() - geo["latitude"].min(),
                   geo["longitude"].max() - geo["longitude"].min())
        zoom = 13 if span < 0.03 else 10 if span < 0.3 else 7.5 if span < 1.5 else 6.3 if span < 4 else 5.4
        common = dict(
            lat="latitude", lon="longitude", size="taille",
            color="statut", color_discrete_map=STATUT_COLORS, hover_name="nom_site",
            hover_data={"client": True, "ville": True, "puissance_totale_kwc": ":.1f",
                        "puissance_onduleur_kva": ":.1f", "stockage_total_kwh": ":.1f",
                        "taille": False, "latitude": False, "longitude": False},
            size_max=34, zoom=zoom, height=620, center={"lat": lat_c, "lon": lon_c})
        if hasattr(px, "scatter_map"):
            fig = px.scatter_map(geo, map_style="open-street-map", **common)
        else:
            fig = px.scatter_mapbox(geo, **common)
            fig.update_layout(mapbox_style="open-street-map")
        fig.update_layout(margin=dict(l=0, r=0, t=0, b=0),
                          font=dict(color=TXT, size=13),
                          legend=dict(orientation="h", yanchor="top", y=.99, x=0,
                                      bgcolor="rgba(255,255,255,.92)", title="",
                                      font=dict(color=TXT, size=13)))
        st.plotly_chart(fig, use_container_width=True)
        cap = f"📍 {len(geo)} installation(s) géolocalisée(s)."
        if sites_sans_gps:
            cap += f" {sites_sans_gps} sans GPS valide."
        st.caption(cap)
    else:
        st.info("Aucune coordonnée GPS valide disponible. Les points apparaîtront dès que les "
                "agents auront renseigné la localisation dans le formulaire Kobo.")

# ==================== 4. RAPPORT PAR INSTALLATION ====================
with tabs[4]:
    st.markdown("<div class='section-title'>Rapport détaillé des puissances installées</div>", unsafe_allow_html=True)
    st.markdown("<div class='section-sub'>Une ligne par installation posée, triable et exportable, "
                "avec le client, la puissance PV, l'onduleur, le stockage et la performance.</div>",
                unsafe_allow_html=True)
    rap_cols = {
        "client": "Client", "nom_site": "Site", "ville": "Ville", "statut": "Statut",
        "puissance_totale_kwc": "PV (kWc)", "puissance_onduleur_kva": "Onduleurs (kVA)",
        "ratio_dimensionnement": "kVA/kWc", "stockage_total_kwh": "Stockage (kWh)",
        "production_reelle_kwh_jour": "Prod. réelle (kWh/j)", "performance_ratio": "Perf. (%)",
    }
    table = df[[c for c in rap_cols if c in df.columns]].rename(columns=rap_cols)
    table = table.sort_values(["Client", "PV (kWc)"], ascending=[True, False])
    st.dataframe(
        table.style.format({
            "PV (kWc)": "{:.1f}", "Onduleurs (kVA)": "{:.1f}", "kVA/kWc": "{:.2f}",
            "Stockage (kWh)": "{:.1f}", "Prod. réelle (kWh/j)": "{:.0f}", "Perf. (%)": "{:.0f}",
        }, na_rep="—").background_gradient(subset=["PV (kWc)"], cmap="Blues"),
        use_container_width=True, hide_index=True, height=430)
    d1, d2 = st.columns(2)
    with d1:
        st.download_button(" Rapport Excel (3 feuilles)", build_excel_report(df),
                           f"rapport_puissances_inevoke_{datetime.now():%Y%m%d}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with d2:
        st.download_button("📥 Tableau (CSV)", table.to_csv(index=False).encode("utf-8-sig"),
                           f"rapport_puissances_inevoke_{datetime.now():%Y%m%d}.csv", "text/csv")

    st.markdown("<div class='section-title'>Classement des installations par puissance</div>", unsafe_allow_html=True)
    top = df.sort_values("puissance_totale_kwc").tail(15)
    fig = px.bar(top, x="puissance_totale_kwc", y="nom_site", orientation="h",
                 text="puissance_totale_kwc", color="client", color_discrete_sequence=PALETTE,
                 hover_data=["ville", "puissance_onduleur_kva", "stockage_total_kwh"])
    fig.update_traces(texttemplate="%{text:.1f} kWc", textposition="outside")
    fig.update_layout(xaxis_title="Puissance PV (kWc)", yaxis_title="")
    st.plotly_chart(styled_fig(fig, 470), use_container_width=True)

# ==================== 5. FICHE CLIENT ====================
with tabs[5]:
    st.markdown("<div class='section-title'>Fiche de puissance par client</div>", unsafe_allow_html=True)
    clients = sorted(df["client"].dropna().unique().tolist())
    if clients:
        sel = st.selectbox("Sélectionner un client", clients)
        cd = df[df["client"] == sel].copy()

        m = st.columns(5)
        vals = [
            ("Sites installés", fmt_num(len(cd), 0)),
            ("Puissance PV", f"{fmt_num(cd['puissance_totale_kwc'].sum(), 1)} kWc"),
            ("Onduleurs", f"{fmt_num(cd['puissance_onduleur_kva'].sum(), 1)} kVA"),
            ("Stockage", f"{fmt_num(cd['stockage_total_kwh'].sum(), 1)} kWh"),
            ("Prod. réelle", f"{fmt_num(cd['production_reelle_kwh_jour'].sum(), 0)} kWh/j"),
        ]
        for col, (l, v) in zip(m, vals):
            col.markdown(f'<div class="mini-metric"><div class="v">{v}</div><div class="l">{l}</div></div>',
                         unsafe_allow_html=True)

        st.markdown("<div class='section-title'>Puissance installée par site</div>", unsafe_allow_html=True)
        cc = cd.sort_values("puissance_totale_kwc")
        fig = px.bar(cc, x="puissance_totale_kwc", y="nom_site", orientation="h",
                     text="puissance_totale_kwc", color="statut", color_discrete_map=STATUT_COLORS)
        fig.update_traces(texttemplate="%{text:.1f} kWc", textposition="outside")
        fig.update_layout(xaxis_title="Puissance PV (kWc)", yaxis_title="")
        st.plotly_chart(styled_fig(fig, max(260, 90 + 42 * len(cc))), use_container_width=True)

        st.markdown("<div class='section-title'>Détail des installations du client</div>", unsafe_allow_html=True)
        det_cols = {
            "installation_id": "ID", "nom_site": "Site", "ville": "Ville", "statut": "Statut",
            "puissance_totale_kwc": "PV (kWc)", "puissance_onduleur_kva": "Onduleurs (kVA)",
            "stockage_total_kwh": "Stockage (kWh)", "date_installation": "Installée le",
            "prochaine_maintenance": "Prochaine maint.",
        }
        det = cd[[c for c in det_cols if c in cd.columns]].rename(columns=det_cols)
        st.dataframe(det, use_container_width=True, hide_index=True)
        st.download_button(f"📥 Exporter les installations de « {sel} » (CSV)",
                           det.to_csv(index=False).encode("utf-8-sig"),
                           f"client_{sel.replace(' ', '_')}.csv", "text/csv")
    else:
        st.info("Aucun client disponible dans la sélection actuelle.")

# ==================== 6. DONNÉES BRUTES ====================
with tabs[6]:
    st.markdown("<div class='section-title'>Données centralisées (installations posées)</div>", unsafe_allow_html=True)
    st.dataframe(df, use_container_width=True, hide_index=True, height=460)
    cc1, cc2 = st.columns(2)
    with cc1:
        st.download_button("📥 Toutes les données (CSV)", df.to_csv(index=False).encode("utf-8-sig"),
                           f"installations_inevoke_{datetime.now():%Y%m%d}.csv", "text/csv")
    with cc2:
        st.download_button(" Rapport de puissances (Excel)", build_excel_report(df),
                           f"rapport_puissances_inevoke_{datetime.now():%Y%m%d}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ==================== 7. COLLECTE & SAISIE (formulaire intégré) ====================
with tabs[7]:
    st.markdown("<div class='section-title'>Collecte & mise à jour des données</div>", unsafe_allow_html=True)
    st.markdown(
        "<div class='section-sub'>Saisissez une nouvelle installation ou mettez à jour une "
        "fiche directement ici. Le formulaire fonctionne aussi hors-ligne : les enregistrements "
        "se synchronisent dès le retour de la connexion.</div>", unsafe_allow_html=True)

    # Bouton d'ouverture dans un nouvel onglet (repli si l'iframe est bloquée par le navigateur)
    st.markdown(
        f"""
        <div style="display:flex;gap:12px;align-items:center;margin:.2rem 0 1rem;">
          <a href="{KOBO_FORM_LINK}" target="_blank" rel="noopener"
             style="background:{ORANGE};color:{INK};font-weight:800;text-decoration:none;
                    padding:.55rem 1.1rem;border-radius:9px;display:inline-block;">
             ➕ Ouvrir le formulaire en plein écran
          </a>
          <span style="color:#7A889B;font-size:12.5px;">
             Après une soumission, cliquez sur « Forcer l'actualisation » dans la barre latérale
             pour recharger les indicateurs.
          </span>
        </div>
        """, unsafe_allow_html=True)

    # Formulaire Enketo embarqué
    try:
        components.iframe(KOBO_FORM_LINK, height=900, scrolling=True)
    except Exception:
        st.warning("Le formulaire ne peut pas s'afficher en intégré ici. Utilisez le bouton "
                   "« Ouvrir le formulaire en plein écran » ci-dessus.")

    st.caption("Si le cadre reste blanc, votre navigateur bloque l'affichage intégré : "
               "ouvrez le formulaire en plein écran via le bouton ci-dessus.")

st.markdown("---")
st.caption("INEVOKE SARL — Suivi des puissances installées par client · Streamlit · Plotly · KoboToolbox")

