import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
from datetime import datetime, timedelta
import openpyxl
import requests
import io

# ══════════════════════════════════════════════════════════════════════════════
#  CONFIG
# ══════════════════════════════════════════════════════════════════════════════
st.set_page_config(page_title="Analyse Délais FAB", page_icon="📊", layout="wide")

st.markdown("""
<style>
    section[data-testid="stSidebar"] { background: #1a3a4a; }
    section[data-testid="stSidebar"] * { color: #e8f4f8 !important; }
    section[data-testid="stSidebar"] .stSelectbox > div > div,
    section[data-testid="stSidebar"] .stMultiSelect > div > div { background: #253e4e; }
    section[data-testid="stSidebar"] hr { border-color: #2d5a6e; }
    .info-table td { padding: 4px 12px; }
    .info-table td:first-child { font-weight:600; color:#1a5276; background:#d6eaf8; }
    h1 { color: #1a5276; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
#  NETWORKDAYS (sans jours fériés)
# ══════════════════════════════════════════════════════════════════════════════
def networkdays(start, end):
    """Nombre de jours ouvrés entre start et end (inclus), comme Excel NETWORKDAYS."""
    if pd.isna(start) or pd.isna(end):
        return np.nan
    s = pd.Timestamp(start).normalize()
    e = pd.Timestamp(end).normalize()
    if s > e:
        return 0
    bdays = pd.bdate_range(s, e)
    return len(bdays)

# ══════════════════════════════════════════════════════════════════════════════
#  CHARGEMENT DES DONNÉES DEPUIS GITHUB
# ══════════════════════════════════════════════════════════════════════════════
@st.cache_data(ttl=3600) # Mise en cache d'une heure pour ne pas re-télécharger à chaque clic
def load_data():
    # L'espace dans le nom du fichier est remplacé par %20 pour être compatible avec les URL
    URL_GITHUB = "https://raw.githubusercontent.com/maximeponceblanc-a11y/D-lais/main/ANALYSE%20DELAIS.xlsx"
    
    try:
        response = requests.get(URL_GITHUB)
        response.raise_for_status() # Vérifie si le téléchargement s'est bien passé
    except requests.exceptions.RequestException as e:
        raise FileNotFoundError(f"Impossible d'accéder au fichier sur GitHub : {e}")

    # On lit les données binaires téléchargées comme si c'était un fichier physique
    file_bytes = io.BytesIO(response.content)
    wb = openpyxl.load_workbook(file_bytes, read_only=True, data_only=True)
    ws = wb.active
    raw = list(ws.iter_rows(values_only=True))
    wb.close()

    headers_raw = list(raw[0])
    seen = {}
    headers = []
    for c in headers_raw:
        k = str(c) if c is not None else "COL"
        if k in seen:
            seen[k] += 1
            headers.append(f"{k}_{seen[k]}")
        else:
            seen[k] = 0
            headers.append(k)

    df = pd.DataFrame(raw[1:], columns=headers)

    DATE_COLS = [
        "DATE COMMANDE", "DATE COMMANDE MATIERE", "DATE COMPLET MATIERE",
        "DATE FICHIER", "DATE VALID MODELE CLIENT", "DATE PAPIER 380 GR",
        "DATE IMPRESSION", "DATE CARTON", "DEPART EN PROD MATIERE",
        "DEPART EN PROD TIRAGES & ACHATS", "DATE DE LIVRAISON INITIALE",
        "DATE DE LIVRAISON REELLE", "DATE DE LA DERNIERE LIVRAISON",
    ]
    for col in DATE_COLS:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    NUMERIC_COLS = [
        "FAB", "QUANTITE", 
        "Délai entre départ matière/ départ impression & achats", 
        "Délai marine: Premier envoi / Première livraison", 
        "Délai marine: Première livraison/ Dernière livraison"
    ]
    for col in NUMERIC_COLS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    F  = "DATE COMMANDE"
    G  = "DATE COMMANDE MATIERE"
    H  = "DATE COMPLET MATIERE"
    I  = "DATE FICHIER"
    J  = "DATE VALID MODELE CLIENT"
    K  = "DATE PAPIER 380 GR"
    L  = "DATE IMPRESSION"
    M  = "DATE CARTON"
    N  = "DEPART EN PROD MATIERE"
    O  = "DEPART EN PROD TIRAGES & ACHATS"
    P  = "DATE DE LIVRAISON INITIALE"
    Q  = "DATE DE LIVRAISON REELLE"
    R  = "DATE DE LA DERNIERE LIVRAISON"

    def nwd(a, b):
        return df.apply(lambda r: networkdays(r.get(a), r.get(b)) - 1
                        if pd.notna(r.get(a)) and pd.notna(r.get(b)) else np.nan, axis=1)

    df["Délai total: Ouverture / Dernière livraison"]              = nwd(F, R)
    df["Délai total: Ouverture / Première livraison"]              = nwd(F, Q)
    df["Ecart délai initial et réel"]                              = nwd(P, Q)
    df["Délai total matière: Ouverture / Départ en prod matière"]  = nwd(F, N)
    df["Délai commande matière: Ouverture / Appel matière"]        = nwd(F, G)
    df["Délai réception matière: Appel matière / Complet matière"] = nwd(G, H)
    df["Délai départ matière: Complet matière / Départ en prod matière"] = nwd(H, N)
    df["Délai total tirages et achats: Ouverture / Départ en prod tirages et achats"] = nwd(F, O)
    df["Délai total Impression : Ouverture / Impression"]          = nwd(F, L)
    df["Délai Papier :Ouverture / Réception papier"]               = nwd(F, K)
    df["Délai Carton :Ouverture / Réception carton"]               = nwd(F, M)
    df["Délai Fichier :Ouverture / fichier définitif"]             = nwd(F, I)
    df["Délai Modèle client : Ouverture / Validation modèle client"] = nwd(F, J)

    def delai_imp_max(r):
        li, lj, ll = r.get(I), r.get(J), r.get(L)
        if pd.isna(ll):
            return np.nan
        base = None
        if pd.notna(li) and pd.notna(lj):
            base = max(li, lj)
        elif pd.notna(li):
            base = li
        elif pd.notna(lj):
            base = lj
        if base is None:
            return np.nan
        return networkdays(base, ll) - 1

    df["Délai Impression : MAX(fichier/modele) / Impression"] = df.apply(delai_imp_max, axis=1)

    def delai_imp_depart(r):
        ll, lo = r.get(L), r.get(O)
        if pd.isna(ll) or pd.isna(lo):
            return np.nan
        return networkdays(ll, lo) - 1

    df["Délai Impression / Départ en prod tirages et achats"] = df.apply(delai_imp_depart, axis=1)

    return df, DATE_COLS

# Configuration des sections incluant les délais placés séquentiellement
SECTIONS = [
    ("Délai total: Ouverture / Dernière livraison",
     "Délai total: Ouverture / Dernière livraison", "#1a7ba6", "DATE DE LA DERNIERE LIVRAISON", 0),
    ("Délai total: Ouverture / Première livraison",
     "Délai total: Ouverture / Première livraison", "#1a7ba6", "DATE DE LIVRAISON REELLE", 0),
    ("Délai entre départ matière/ départ impression & achats",
     "Délai entre départ matière/ départ impression & achats", "#e59866", "DEPART EN PROD TIRAGES & ACHATS", None),
    ("Délai marine: Premier envoi / Première livraison",
     "Délai marine: Premier envoi / Première livraison", "#d35400", "DATE DE LIVRAISON REELLE", None),
    ("Délai marine: Première livraison/ Dernière livraison",
     "Délai marine: Première livraison/ Dernière livraison", "#566573", "DATE DE LA DERNIERE LIVRAISON", None),
    ("Délai total matière: Ouverture / Départ en prod matière",
     "Délai total matière: Ouverture / Départ en prod matière", "#e8a090", "DEPART EN PROD MATIERE", 0),
    ("Délai commande matière: Ouverture / Appel matière",
     "Délai commande matière: Ouverture / Appel matière", "#f4b8a8", "DATE COMMANDE MATIERE", 0),
    ("Délai réception matière: Appel matière / Complet matière",
     "Délai réception matière: Appel matière / Complet matière", "#f4b8a8", "DATE COMPLET MATIERE", None),
    ("Délai départ matière: Complet matière / Départ en prod matière",
     "Délai départ matière: Complet matière / Départ en prod matière", "#f4b8a8", "DEPART EN PROD MATIERE", None),
    ("Délai total tirages & achats: Ouverture / Départ en prod",
     "Délai total tirages et achats: Ouverture / Départ en prod tirages et achats", "#2d7a3a", "DEPART EN PROD TIRAGES & ACHATS", 0),
    ("Délai total Impression: Ouverture / Impression",
     "Délai total Impression : Ouverture / Impression", "#3a9147", "DATE IMPRESSION", 0),
    ("Délai Papier: Ouverture / Réception papier",
     "Délai Papier :Ouverture / Réception papier", "#4aab59", "DATE PAPIER 380 GR", 0),
    ("Délai Carton: Ouverture / Réception carton",
     "Délai Carton :Ouverture / Réception carton", "#4aab59", "DATE CARTON", 0),
    ("Délai Fichier: Ouverture / Fichier définitif",
     "Délai Fichier :Ouverture / fichier définitif", "#4aab59", "DATE FICHIER", 0),
    ("Délai Modèle client: Ouverture / Validation",
     "Délai Modèle client : Ouverture / Validation modèle client", "#4aab59", "DATE VALID MODELE CLIENT", 0),
    ("Délai Impression: MAX(fichier/modele) / Impression",
     "Délai Impression : MAX(fichier/modele) / Impression", "#196f3d", "DATE IMPRESSION", None),
    ("Délai Impression / Départ en prod tirages et achats",
     "Délai Impression / Départ en prod tirages et achats", "#0d4a26", "DEPART EN PROD TIRAGES & ACHATS", None),
]

# ══════════════════════════════════════════════════════════════════════════════
#  HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def fmt_date(val):
    if pd.isna(val) or val is None:
        return "—"
    try:
        return pd.Timestamp(val).strftime("%d/%m/%Y")
    except Exception:
        return "—"

def fmt_val(val):
    if pd.isna(val) or val is None:
        return "—"
    if isinstance(val, (int, float, np.integer, np.floating)):
        return str(int(val))
    return str(val)

def safe_float(val):
    try:
        v = float(val)
        return np.nan if np.isnan(v) else v
    except (TypeError, ValueError):
        return np.nan

def get_mean(series):
    vals = series.dropna()
    vals = vals[vals >= 0]
    return vals.mean() if len(vals) > 0 else np.nan

# ══════════════════════════════════════════════════════════════════════════════
#  CHARGEMENT
# ══════════════════════════════════════════════════════════════════════════════
try:
    df, DATE_COLS = load_data()
except FileNotFoundError as e:
    st.error(f"❌ {e}")
    st.stop()
except Exception as e:
    st.error(f"❌ Une erreur inattendue est survenue lors du chargement : {e}")
    st.stop()

ALL_DATE_COLS = [c for c in DATE_COLS if c in df.columns]
DELAI_COLS = [s[1] for s in SECTIONS]

# ══════════════════════════════════════════════════════════════════════════════
#  BARRE LATÉRALE
# ══════════════════════════════════════════════════════════════════════════════
range_sliders = {}
with st.sidebar:
    st.markdown("## 📊 Navigation")
    vue = st.radio("Vue", ["📁 Vue par dossier", "📈 Délais moyens"], label_visibility="collapsed")
    st.markdown("---")

    if vue == "📁 Vue par dossier":
        st.markdown("### 🔍 Filtre dossier")
        fab_list = sorted(df["FAB"].dropna().astype(int).unique().tolist())
        fab_choice = st.selectbox("Numéro FAB", options=fab_list)
    else:
        st.markdown("### 🎚️ Filtrer par plages de délais")
        st.caption("Sélectionnez les valeurs minimales et maximales acceptées pour chaque délai (en j.o.).")
        
        for label, col, color, _, _ in SECTIONS:
            if col in df.columns:
                vals = df[col].dropna()
                if len(vals) > 0:
                    min_val = int(vals.min())
                    max_val = int(vals.max())
                    if min_val == max_val:
                        max_val += 1
                    lo, hi = st.slider(
                        label[:45] + ("…" if len(label) > 45 else ""),
                        min_value=min_val,
                        max_value=max_val,
                        value=(min_val, max_val),
                        key=f"sl_{col}",
                        help=f"Garder uniquement les valeurs entre {min_val} et {max_val} j.o.",
                    )
                    range_sliders[col] = (lo, hi, min_val, max_val)

        st.markdown("---")
        st.markdown("### 🔎 Filtres données")

        clients = sorted(df["CLIENT"].dropna().unique().tolist())
        client_filter = st.multiselect("Client(s)", options=clients, default=[])

        q_min_all = int(df["QUANTITE"].min(skipna=True)) if pd.notna(df["QUANTITE"].min()) else 0
        q_max_all = int(df["QUANTITE"].max(skipna=True)) if pd.notna(df["QUANTITE"].max()) else 100
        if q_min_all < q_max_all:
            q_range = st.slider("Quantité", min_value=q_min_all, max_value=q_max_all, value=(q_min_all, q_max_all))
        else:
            q_range = (q_min_all, q_max_all)

        dates_valid = df["DATE COMMANDE"].dropna()
        if len(dates_valid):
            d_min = dates_valid.min().date()
            d_max = dates_valid.max().date()
            date_range = st.date_input("Date de commande", value=(d_min, d_max), min_value=d_min, max_value=d_max)
        else:
            date_range = None

# ══════════════════════════════════════════════════════════════════════════════
#  APPLICATION DES FILTRES
# ══════════════════════════════════════════════════════════════════════════════
def apply_filters(data):
    d = data.copy()
    if client_filter:
        d = d[d["CLIENT"].isin(client_filter)]
    if "QUANTITE" in d.columns:
        d = d[d["QUANTITE"].between(q_range[0], q_range[1], inclusive="both") | d["QUANTITE"].isna()]
    if date_range and len(date_range) == 2:
        lo_d = pd.Timestamp(date_range[0])
        hi_d = pd.Timestamp(date_range[1])
        d = d[(d["DATE COMMANDE"] >= lo_d) & (d["DATE COMMANDE"] <= hi_d) | d["DATE COMMANDE"].isna()]
        
    for col, (lo, hi, min_val, max_val) in range_sliders.items():
        if col in d.columns:
            if lo == min_val and hi == max_val:
                d = d[d[col].between(lo, hi, inclusive="both") | d[col].isna()]
            else:
                d = d[d[col].between(lo, hi, inclusive="both")]
    return d

# ══════════════════════════════════════════════════════════════════════════════
#  TITRE
# ══════════════════════════════════════════════════════════════════════════════
st.title("📊 Tableau de bord — Analyse des délais de fabrication")
st.markdown("---")

# ══════════════════════════════════════════════════════════════════════════════
#  VUE 1 : PAR DOSSIER
# ══════════════════════════════════════════════════════════════════════════════
if vue == "📁 Vue par dossier":
    rows = df[df["FAB"] == fab_choice]
    if rows.empty:
        st.warning("Aucune donnée pour ce FAB.")
        st.stop()

    if len(rows) > 1:
        prod_list = rows["NOM PRODUIT"].fillna("—").tolist()
        prod_choice = st.selectbox("📦 Produit", options=prod_list)
        row = rows[rows["NOM PRODUIT"] == prod_choice].iloc[0]
    else:
        row = rows.iloc[0]

    col_info, col_metrics = st.columns([2, 3])
    with col_info:
        st.markdown("#### 📋 Informations commande")
        info_data = {
            "FAB": fmt_val(row.get("FAB")),
            "CLIENT": fmt_val(row.get("CLIENT")),
            "NOM PRODUIT": fmt_val(row.get("NOM PRODUIT")),
            "QUANTITÉ": fmt_val(row.get("QUANTITE")),
            "TYPE NUANCIER": fmt_val(row.get("TYPE NUANCIER")),
            "NB DÉPART": fmt_val(row.get("NB DEPART ")),
            "TYPE BLOCAGE": fmt_val(row.get("TYPE BLOCAGE")),
            "INFO": fmt_val(row.get("INFO")),
        }
        html_rows = "".join(f"<tr><td>{k}</td><td>{v}</td></tr>" for k, v in info_data.items() if v != "—")
        st.markdown(f'<table class="info-table">{html_rows}</table>', unsafe_allow_html=True)

    with col_metrics:
        st.markdown("#### 📅 Métriques clés")
        m1, m2, m3 = st.columns(3)
        m4, m5, _ = st.columns(3)
        m1.metric("Date commande", fmt_date(row.get("DATE COMMANDE")))
        m2.metric("Dernière livraison", fmt_date(row.get("DATE DE LA DERNIERE LIVRAISON")))
        m3.metric("Livraison réelle", fmt_date(row.get("DATE DE LIVRAISON REELLE")))
        m4.metric("Délai total (j.o.)", fmt_val(row.get("Délai total: Ouverture / Dernière livraison")))
        m5.metric("Écart initial/réel", fmt_val(row.get("Ecart délai initial et réel")))

    st.markdown("---")
    st.subheader(f"📐 Diagramme des délais — FAB {int(fab_choice)}")

    def get_v(col):
        return safe_float(row.get(col))

    d_cmd_mat = get_v("Délai commande matière: Ouverture / Appel matière")
    d_rec_mat = get_v("Délai réception matière: Appel matière / Complet matière")
    d_fichier  = get_v("Délai Fichier :Ouverture / fichier définitif")
    d_modele   = get_v("Délai Modèle client : Ouverture / Validation modèle client")
    d_imp_sub  = get_v("Délai Impression : MAX(fichier/modele) / Impression")
    max_f_m    = max(v for v in [d_fichier, d_modele] if not np.isnan(v)) if any(not np.isnan(v) for v in [d_fichier, d_modele]) else 0

    d_tot_mat = get_v("Délai total matière: Ouverture / Départ en prod matière")
    d_tot_tir = get_v("Délai total tirages et achats: Ouverture / Départ en prod tirages et achats")
    valid_tot = [v for v in [d_tot_mat, d_tot_tir] if not np.isnan(v)]
    
    min_mat_tir = min(valid_tot) if valid_tot else 0
    max_mat_tir = max(valid_tot) if valid_tot else 0

    # Pour un dossier unique, la valeur absolue reste nécessaire pour l'affichage physique de l'écart
    d_entre_dept = get_v("Délai entre départ matière/ départ impression & achats")
    d_tot_prem_liv = get_v("Délai total: Ouverture / Première livraison")

    def base_for(label):
        if "Délai entre départ matière/ départ impression & achats" in label:
            return min_mat_tir
        if "Délai marine: Premier envoi / Première livraison" in label:
            return max_mat_tir
        if "Délai marine: Première livraison/ Dernière livraison" in label:
            return d_tot_prem_liv if not np.isnan(d_tot_prem_liv) else 0
        if "Appel matière / Complet" in label: return d_cmd_mat if not np.isnan(d_cmd_mat) else 0
        if "Complet matière / Départ" in label: return (d_cmd_mat if not np.isnan(d_cmd_mat) else 0) + (d_rec_mat if not np.isnan(d_rec_mat) else 0)
        if "MAX(fichier/modele) / Impression" in label: return max_f_m
        if "Impression / Départ en prod" in label:
            d_imp_tot = get_v("Délai total Impression : Ouverture / Impression")
            return d_imp_tot if not np.isnan(d_imp_tot) else max_f_m + (d_imp_sub if not np.isnan(d_imp_sub) else 0)
        return 0

    labels, values, colors, texts, bases = [], [], [], [], []
    for label, col, color, _, _ in SECTIONS:
        val = get_v(col)
        if not np.isnan(val):
            base = base_for(label)
            labels.append(label)
            values.append(val)
            colors.append(color)
            texts.append(f"<b>{int(val)} j.o.</b>")
            bases.append(base if not np.isnan(base) else 0)

    if labels:
        fig = go.Figure()
        for lbl, val, col, txt, base in zip(labels, values, colors, texts, bases):
            fig.add_trace(go.Bar(
                name=lbl, y=[lbl], x=[val], base=[base], orientation="h",
                marker_color=col, marker_line_color="white", marker_line_width=1,
                text=txt, textposition="inside", insidetextanchor="middle",
                textfont=dict(color="white", size=11, family="Arial"),
                hovertemplate=f"<b>{lbl}</b><br>Délai : %{{x}} j.o.<extra></extra>",
                showlegend=False,
            ))
        max_val = max(b + v for b, v in zip(bases, values))
        fig.update_layout(
            xaxis=dict(title="Jours ouvrés", showgrid=True, gridcolor="#e0e0e0", range=[0, max_val * 1.12]),
            yaxis=dict(autorange="reversed", tickfont=dict(size=10)),
            height=60 + len(labels) * 50, margin=dict(l=380, r=20, t=30, b=40),
            plot_bgcolor="white", paper_bgcolor="#f8f9fa", bargap=0.25,
        )
        st.plotly_chart(fig, use_container_width=True, key="chart_dossier")

    st.markdown("---")
    st.subheader("📅 Dates clés")
    DATE_LABELS_MAP = {
        "DATE COMMANDE": "Date commande", "DATE COMMANDE MATIERE": "Date commande matière",
        "DATE COMPLET MATIERE": "Date complet matière", "DATE FICHIER": "Date fichier définitif",
        "DATE VALID MODELE CLIENT": "Date validation modèle client", "DATE PAPIER 380 GR": "Date papier 380 gr",
        "DATE IMPRESSION": "Date impression", "DATE CARTON": "Date carton",
        "DEPART EN PROD MATIERE": "Départ en prod matière", "DEPART EN PROD TIRAGES & ACHATS": "Départ en prod tirages & achats",
        "DATE DE LIVRAISON INITIALE": "Date livraison initiale", "DATE DE LIVRAISON REELLE": "Date livraison réelle",
        "DATE DE LA DERNIERE LIVRAISON": "Date dernière livraison",
    }
    date_rows = [{"Étape": lbl, "Date": fmt_date(row.get(col))} for col, lbl in DATE_LABELS_MAP.items() if col in df.columns]
    st.dataframe(pd.DataFrame(date_rows), use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════════════════
#  VUE 2 : DÉLAIS MOYENS (AVEC LA CORRECTION DE L'ÉCART ALGEBRIQUE MOYEN EN J.O.)
# ══════════════════════════════════════════════════════════════════════════════
else:
    df_f = apply_filters(df)
    if df_f.empty:
        st.warning("⚠️ Aucun dossier ne correspond à la combinaison de vos critères.")
        st.stop()

    st.markdown("#### 📈 Analyse de performance globale — Dossier moyen fictif")
    st.caption(f"Données affichées : **{len(df_f)}** dossiers sur {len(df)} au total.")

    c1, c2, c3, c4 = st.columns(4)
    avg_total = get_mean(df_f["Délai total: Ouverture / Dernière livraison"])
    avg_ecart = get_mean(df_f["Ecart délai initial et réel"])
    avg_mat   = get_mean(df_f["Délai total matière: Ouverture / Départ en prod matière"])

    c1.metric("📁 Dossiers filtrés", len(df_f))
    c2.metric("⏱ Moy. délai total", f"{avg_total:.1f} j.o." if pd.notna(avg_total) else "—")
    c3.metric("⚖️ Moy. écart init/réel", f"{avg_ecart:.1f} j.o." if pd.notna(avg_ecart) else "—")
    c4.metric("🪵 Moy. délai matière", f"{avg_mat:.1f} j.o." if pd.notna(avg_mat) else "—")

    st.markdown("---")
    st.subheader("📐 Diagramme des délais moyens")

    def get_avg(col):
        return get_mean(df_f[col]) if col in df_f.columns else np.nan

    avg_cmd_mat = get_avg("Délai commande matière: Ouverture / Appel matière")
    avg_rec_mat = get_avg("Délai réception matière: Appel matière / Complet matière")
    avg_fichier  = get_avg("Délai Fichier :Ouverture / fichier définitif")
    avg_modele   = get_avg("Délai Modèle client : Ouverture / Validation modèle client")
    avg_imp_sub  = get_avg("Délai Impression : MAX(fichier/modele) / Impression")
    avg_imp_tot  = get_avg("Délai total Impression : Ouverture / Impression")

    valid_fm = [v for v in [avg_fichier, avg_modele] if not np.isnan(v)]
    max_f_m_avg = max(valid_fm) if valid_fm else 0

    avg_tot_mat = get_avg("Délai total matière: Ouverture / Départ en prod matière")
    avg_tot_tir = get_avg("Délai total tirages et achats: Ouverture / Départ en prod tirages et achats")
    valid_tot_avg = [v for v in [avg_tot_mat, avg_tot_tir] if not np.isnan(v)]
    
    min_mat_tir_avg = min(valid_tot_avg) if valid_tot_avg else 0
    max_mat_tir_avg = max(valid_tot_avg) if valid_tot_avg else 0

    # ══════════════════════════════════════════════════════════════════════════
    # CORRECTION CRITIQUE : CALCUL DE L'ÉCART MOYEN COMPENSÉ EN JOURS OUVRÉS
    # ══════════════════════════════════════════════════════════════════════════
    # On calcule l'écart algébrique réel en jours ouvrés pour chaque dossier filtré
    ecarts_individuels_jo = df_f["Délai total matière: Ouverture / Départ en prod matière"] - df_f["Délai total tirages et achats: Ouverture / Départ en prod tirages et achats"]
    
    # On fait la moyenne arithmétique simple (les + et les - s'annulent/se compensent ici)
    moyenne_des_ecarts_jo = ecarts_individuels_jo.dropna().mean()
    
    # On prend la valeur absolue à la toute fin pour l'affichage de la taille du bloc
    avg_entre_dept = abs(moyenne_des_ecarts_jo) if pd.notna(moyenne_des_ecarts_jo) else np.nan
    # ══════════════════════════════════════════════════════════════════════════

    avg_tot_prem_liv = get_avg("Délai total: Ouverture / Première livraison")

    def base_avg_for(label):
        if "Délai entre départ matière/ départ impression & achats" in label:
            return min_mat_tir_avg
        if "Délai marine: Premier envoi / Première livraison" in label:
            return max_mat_tir_avg
        if "Délai marine: Première livraison/ Dernière livraison" in label:
            return avg_tot_prem_liv if not np.isnan(avg_tot_prem_liv) else 0
        if "Appel matière / Complet" in label: return avg_cmd_mat if not np.isnan(avg_cmd_mat) else 0
        if "Complet matière / Départ" in label: return (avg_cmd_mat if not np.isnan(avg_cmd_mat) else 0) + (avg_rec_mat if not np.isnan(avg_rec_mat) else 0)
        if "MAX(fichier/modele) / Impression" in label: return max_f_m_avg
        if "Impression / Départ en prod" in label: return avg_imp_tot if not np.isnan(avg_imp_tot) else max_f_m_avg + (avg_imp_sub if not np.isnan(avg_imp_sub) else 0)
        return 0

    labels_a, values_a, colors_a, texts_a, bases_a = [], [], [], [], []
    for label, col, color, _, _ in SECTIONS:
        # Si c'est notre colonne modifiée, on force notre nouvelle valeur calculée
        if col == "Délai entre départ matière/ départ impression & achats":
            val = avg_entre_dept
        else:
            val = get_avg(col)
            
        if pd.notna(val):
            base = base_avg_for(label)
            labels_a.append(label)
            values_a.append(float(val))
            colors_a.append(color)
            texts_a.append(f"<b>{val:.1f} j.o.</b>")
            bases_a.append(float(base) if not np.isnan(base) else 0)

    if labels_a:
        fig_avg = go.Figure()
        for lbl, val, col, txt, base in zip(labels_a, values_a, colors_a, texts_a, bases_a):
            fig_avg.add_trace(go.Bar(
                name=lbl, y=[lbl], x=[val], base=[base], orientation="h",
                marker_color=col, marker_line_color="white", marker_line_width=1,
                text=txt, textposition="inside", insidetextanchor="middle",
                textfont=dict(color="white", size=11, family="Arial"),
                hovertemplate=f"<b>{lbl}</b><br>Moyenne : %{{x:.1f}} j.o.<extra></extra>",
                showlegend=False,
            ))
        max_val_a = max(b + v for b, v in zip(bases_a, values_a))
        fig_avg.update_layout(
            xaxis=dict(title="Jours ouvrés (Moyenne)", showgrid=True, gridcolor="#e0e0e0", range=[0, max_val_a * 1.12]),
            yaxis=dict(autorange="reversed", tickfont=dict(size=10)),
            height=60 + len(labels_a) * 50, margin=dict(l=380, r=20, t=30, b=40),
            plot_bgcolor="white", paper_bgcolor="#f8f9fa", bargap=0.25,
        )
        st.plotly_chart(fig_avg, use_container_width=True, key="chart_avg")

    st.markdown("---")
    st.subheader("🔵 Nuages de points — Délais par date d'ouverture")

    DELAI_DATE_MAP = {
        "Délai total: Ouverture / Dernière livraison":              "DATE DE LA DERNIERE LIVRAISON",
        "Délai total: Ouverture / Première livraison":              "DATE DE LIVRAISON REELLE",
        "Délai entre départ matière/ départ impression & achats":   "DEPART EN PROD TIRAGES & ACHATS",
        "Délai marine: Premier envoi / Première livraison":         "DATE DE LIVRAISON REELLE",
        "Délai marine: Première livraison/ Dernière livraison":     "DATE DE LA DERNIERE LIVRAISON",
        "Délai total matière: Ouverture / Départ en prod matière":  "DEPART EN PROD MATIERE",
        "Délai commande matière: Ouverture / Appel matière":        "DATE COMMANDE MATIERE",
        "Délai réception matière: Appel matière / Complet matière": "DATE COMPLET MATIERE",
        "Délai départ matière: Complet matière / Départ en prod matière": "DEPART EN PROD MATIERE",
        "Délai total tirages et achats: Ouverture / Départ en prod tirages et achats": "DEPART EN PROD TIRAGES & ACHATS",
        "Délai total Impression : Ouverture / Impression":          "DATE IMPRESSION",
        "Délai Papier :Ouverture / Réception papier":               "DATE PAPIER 380 GR",
        "Délai Carton :Ouverture / Réception carton":               "DATE CARTON",
        "Délai Fichier :Ouverture / fichier définitif":             "DATE FICHIER",
        "Délai Modèle client : Ouverture / Validation modèle client": "DATE VALID MODELE CLIENT",
        "Délai Impression : MAX(fichier/modele) / Impression":      "DATE IMPRESSION",
        "Délai Impression / Départ en prod tirages et achats":      "DEPART EN PROD TIRAGES & ACHATS",
    }

    q_vals = df_f["QUANTITE"].dropna()
    q_min_v = q_vals.min() if len(q_vals) else 1
    q_max_v = q_vals.max() if len(q_vals) else 1
    q_range_v = q_max_v - q_min_v if q_max_v > q_min_v else 1

    def bubble_size(q):
        if pd.isna(q): return 8
        return 8 + 30 * ((q - q_min_v) / q_range_v)

    section_pairs = list(zip(SECTIONS[::2], SECTIONS[1::2] + [None] * (len(SECTIONS) % 2)))

    for left, right in section_pairs:
        cols = st.columns(2)
        for col_idx, section in enumerate([left, right]):
            if section is None: continue
            label, delai_col, color, _, _ = section
            date_col = DELAI_DATE_MAP.get(delai_col)

            if delai_col not in df_f.columns:
                continue

            sub = df_f[["FAB", "CLIENT", "NOM PRODUIT", "QUANTITE", "DATE COMMANDE", delai_col] + ([date_col] if date_col and date_col in df_f.columns else [])].dropna(subset=["DATE COMMANDE", delai_col])
            if sub.empty:
                with cols[col_idx]:
                    st.caption(f"🔍 *{label}* — aucun point")
                continue

            sub = sub.copy()
            sub["_size"] = sub["QUANTITE"].apply(bubble_size)
            sub["_date_assoc"] = sub[date_col].apply(fmt_date) if date_col and date_col in sub.columns else "—"
            sub["_hover_fab"]  = sub["FAB"].apply(lambda x: str(int(x)) if pd.notna(x) else "—")
            sub["_hover_q"]    = sub["QUANTITE"].apply(lambda x: str(int(x)) if pd.notna(x) else "—")

            fig_sc = go.Figure()
            fig_sc.add_trace(go.Scatter(
                x=sub["DATE COMMANDE"], y=sub[delai_col], mode="markers",
                marker=dict(size=sub["_size"], color=color, opacity=0.75, line=dict(color="white", width=1)),
                customdata=np.stack([sub["_hover_fab"], sub["CLIENT"].fillna("—"), sub["NOM PRODUIT"].fillna("—"), sub["_hover_q"], sub["DATE COMMANDE"].apply(fmt_date), sub["_date_assoc"]], axis=1),
                hovertemplate=("<b>FAB %{customdata[0]}</b><br>Client : %{customdata[1]}<br>Produit : %{customdata[2]}<br>Quantité : %{customdata[3]}<br>Date commande : %{customdata[4]}<br>Date associée : %{customdata[5]}<br><b>Délai : %{y:.0f} j.o.</b><extra></extra>"),
                showlegend=False,
            ))

            # Affichage de la ligne de moyenne modifiée également pour le nuage de points dédié
            if delai_col == "Délai entre départ matière/ départ impression & achats":
                moy = avg_entre_dept
            else:
                moy = get_mean(sub[delai_col])
                
            if pd.notna(moy):
                fig_sc.add_hline(y=moy, line_dash="dash", line_color=color, annotation_text=f"Moy. {moy:.1f} j.o.", annotation_position="top right", annotation_font_color=color)

            fig_sc.update_layout(
                title=dict(text=f"<b>{label}</b>", font=dict(size=12, color="#1a5276"), x=0),
                xaxis=dict(title="Date commande", showgrid=True, gridcolor="#ececec", tickformat="%b %Y"),
                yaxis=dict(title="Jours ouvrés", showgrid=True, gridcolor="#ececec"),
                height=320, margin=dict(l=50, r=20, t=45, b=50), plot_bgcolor="white", paper_bgcolor="#f8f9fa",
            )
            with cols[col_idx]:
                st.plotly_chart(fig_sc, use_container_width=True, key=f"scatter_{delai_col}")

    st.markdown("---")
    with st.expander("🗂 Tableau de données complet (dossiers filtrés)", expanded=False):
        base_cols = ["FAB", "CLIENT", "NOM PRODUIT", "QUANTITE", "TYPE NUANCIER", "NB DEPART ", "TYPE BLOCAGE", "INFO"]
        all_date_nice = ["DATE COMMANDE", "DATE COMMANDE MATIERE", "DATE COMPLET MATIERE", "DATE FICHIER", "DATE VALID MODELE CLIENT", "DATE PAPIER 380 GR", "DATE IMPRESSION", "DATE CARTON", "DEPART EN PROD MATIERE", "DEPART EN PROD TIRAGES & ACHATS", "DATE DE LIVRAISON INITIALE", "DATE DE LIVRAISON REELLE", "DATE DE LA DERNIERE LIVRAISON"]
        delai_display = [
            "Délai total: Ouverture / Dernière livraison", 
            "Délai total: Ouverture / Première livraison", 
            "Délai entre départ matière/ départ impression & achats",
            "Délai marine: Premier envoi / Première livraison",
            "Délai marine: Première livraison/ Dernière livraison",
            "Ecart délai initial et réel", 
            "Délai total matière: Ouverture / Départ en prod matière", 
            "Délai total tirages et achats: Ouverture / Départ en prod tirages et achats", 
            "Délai total Impression : Ouverture / Impression"
        ]
        all_cols = ([c for c in base_cols if c in df_f.columns] + [c for c in all_date_nice if c in df_f.columns] + [c for c in delai_display if c in df_f.columns])
        df_show = df_f[all_cols].copy()
        for col in all_date_nice:
            if col in df_show.columns:
                df_show[col] = df_show[col].apply(fmt_date)
        if "FAB" in df_show.columns:
            df_show["FAB"] = df_show["FAB"].apply(lambda x: str(int(x)) if pd.notna(x) else "—")
        if "QUANTITE" in df_show.columns:
            df_show["QUANTITE"] = df_show["QUANTITE"].apply(lambda x: str(int(x)) if pd.notna(x) else "—")
        st.dataframe(df_show, use_container_width=True, hide_index=True)

st.markdown("<div style='text-align:center;color:#aaa;font-size:11px;margin-top:2rem;'>Tableau de bord Analyse Délais — données synchronisées avec GitHub</div>", unsafe_allow_html=True)
