import streamlit as st
import pandas as pd
import re
from io import BytesIO

# Provo a importare reportlab solo se disponibile
try:
    import requests
    from reportlab.platypus import (
        SimpleDocTemplate,
        Paragraph,
        Spacer,
        Table,
        TableStyle,
        Image as RLImage,
    )
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    REPORTLAB_AVAILABLE = True
except ModuleNotFoundError:
    REPORTLAB_AVAILABLE = False

from pandas.errors import EmptyDataError, ParserError

# -------------------------------------------------------
# CONFIGURAZIONE BASE, COLORI E LOGO
# -------------------------------------------------------
st.set_page_config(layout="wide", page_title="Dashboard Interventi – Prov. Sulcis Iglesiente")

LOGO_URL = "https://provincia-sulcis-iglesiente-api.cloud.municipiumapp.it/s3/150x150/s3/20243/sito/stemma.jpg"

PRIMARY_HEX = "#6BE600"
PRIMARY_LIGHT = "#A8FF66"
PRIMARY_EXTRA_LIGHT = "#E8FFE0"

st.markdown(
    f"""
    <style>
    [data-testid="stAppViewContainer"] {{
        background: linear-gradient(135deg, {PRIMARY_EXTRA_LIGHT} 0%, #FFFFFF 40%, {PRIMARY_EXTRA_LIGHT} 100%);
    }}

    .sulcis-main-header {{
        display:flex;
        align-items:center;
        gap:1rem;
        background: linear-gradient(90deg, {PRIMARY_HEX} 0%, {PRIMARY_LIGHT} 50%, {PRIMARY_EXTRA_LIGHT} 100%);
        padding: 0.9rem 1.3rem;
        border-radius: 0.75rem;
        margin-bottom: 1.2rem;
        color: #1E2A10;
        font-weight: 600;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
    }}
    .sulcis-main-header-text small {{
        display:block;
        font-weight:400;
        opacity:0.9;
        margin-top:0.2rem;
    }}

    .sulcis-card {{
        background: linear-gradient(135deg, #FFFFFF 0%, {PRIMARY_EXTRA_LIGHT} 100%);
        border-radius: 0.75rem;
        padding: 0.9rem 1.1rem;
        margin-bottom: 1rem;
        border: 1px solid rgba(107,230,0,0.25);
    }}

    .sulcis-section-title {{
        font-weight: 600;
        color: #1E2A10;
        margin-bottom: 0.3rem;
    }}

    h1, h2, h3 {{
        margin-top: 0.4rem;
        margin-bottom: 0.4rem;
    }}
    </style>
    """,
    unsafe_allow_html=True,
)

# Header con logo + testo
col_logo, col_title = st.columns([1, 6])
with col_logo:
    st.image(LOGO_URL, width=70)
with col_title:
    st.markdown(
        """
        <div class="sulcis-main-header">
            <div class="sulcis-main-header-text">
                Dashboard Interventi Istituti Scolastici<br/>
                <small>Provincia del Sulcis Iglesiente – interventi e manutenzioni sugli istituti scolastici</small>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

# -------------------------------------------------------
# HELPER: DF per riepiloghi economici
# regola: stessa (codice, determina_norm, importo_stanziato) = 1 volta
# -------------------------------------------------------
def df_riepilogo(df_source: pd.DataFrame) -> pd.DataFrame:
    if not {"codice", "determina_norm", "importo_stanziato"}.issubset(df_source.columns):
        return df_source
    return df_source.drop_duplicates(subset=["codice", "determina_norm", "importo_stanziato"])

# -------------------------------------------------------
# LETTURA CSV
# -------------------------------------------------------
def load_uploaded_csv(uploaded_file, nome_log="file"):
    if uploaded_file is None:
        st.error(f"Nessun file caricato per {nome_log}.")
        return pd.DataFrame()
    try:
        uploaded_file.seek(0)
        df = pd.read_csv(uploaded_file, sep=";", encoding="utf-8", engine="python")
        if df.empty:
            st.error(f"{nome_log}: il file è vuoto.")
        return df
    except (EmptyDataError, ParserError) as e:
        st.error(f"{nome_log}: errore parsing (UTF-8): {e}")
        return pd.DataFrame()
    except UnicodeDecodeError:
        try:
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file, sep=";", encoding="cp1252", engine="python")
            if df.empty:
                st.error(f"{nome_log}: il file è vuoto (cp1252).")
            return df
        except Exception as e2:
            st.error(f"{nome_log}: fallback cp1252 fallisce: {e2}")
            return pd.DataFrame()
    except Exception as e:
        st.error(f"{nome_log}: errore imprevisto: {e}")
        return pd.DataFrame()

# -------------------------------------------------------
# UPLOAD FILE
# -------------------------------------------------------
st.sidebar.subheader("📂 Caricamento dati")
file_istituti = st.sidebar.file_uploader("File ISTITUTI (SCU_Istituti-ELE_ISTITUTI-2.csv)", type="csv")
file_interventi = st.sidebar.file_uploader("File INTERVENTI (SCU_Istituti-ELE_CMPLSS.csv)", type="csv")

if not file_istituti or not file_interventi:
    st.info("Carica entrambi i file CSV per visualizzare la dashboard.")
    st.stop()

istituti = load_uploaded_csv(file_istituti, "ISTITUTI")
interventi = load_uploaded_csv(file_interventi, "INTERVENTI")

if istituti.empty or interventi.empty:
    st.stop()

istituti.columns = istituti.columns.str.strip()
interventi.columns = interventi.columns.str.strip()

# -------------------------------------------------------
# RINOMINO COLONNE
# -------------------------------------------------------
istituti = istituti.rename(columns={
    "Denominazione Immobile": "nome_istituto",
    "Localizzazione immobile": "indirizzo",
    "Comune": "comune",
    "comune": "comune",
})

interventi = interventi.rename(columns={
    "Nome Istituto": "nome_istituto_descr",
    "Denominazione intervento": "denominazione_intervento",
    "Determina": "determina",
    "Manutenzioni": "manutenzioni",
    "Tipologia di intervento": "tipologia_intervento",
    "RUP": "rup",
    "importo stanziato": "importo_stanziato",
})

if "tipologia_intervento" not in interventi.columns:
    interventi["tipologia_intervento"] = "Non specificata"

# -------------------------------------------------------
# CONTROLLI COLONNE
# -------------------------------------------------------
for col in ["codice", "nome_istituto", "comune"]:
    if col not in istituti.columns:
        st.error(f"File ISTITUTI: colonna '{col}' mancante. Colonne trovate: {list(istituti.columns)}")
        st.stop()

for col in ["codice", "denominazione_intervento", "determina", "manutenzioni", "tipologia_intervento"]:
    if col not in interventi.columns:
        st.error(f"File INTERVENTI: colonna '{col}' mancante. Colonne trovate: {list(interventi.columns)}")
        st.stop()

# -------------------------------------------------------
# JOIN SU 'codice'
# -------------------------------------------------------
istituti["codice"] = istituti["codice"].astype(str).str.strip()
interventi["codice"] = interventi["codice"].astype(str).str.strip()

df = interventi.merge(
    istituti[["codice", "nome_istituto", "comune", "indirizzo"]],
    on="codice",
    how="left",
)

# -------------------------------------------------------
# NORMALIZZAZIONE / FLAG
# -------------------------------------------------------
df["determina_norm"] = df["determina"].astype(str).str.strip().str.lower()
df["manut_flag"] = df["manutenzioni"].astype(str).str.lower().eq("vero")
df["tipologia_intervento"] = df["tipologia_intervento"].astype(str).str.strip()

# -------------------------------------------------------
# PULIZIA IMPORTI
# -------------------------------------------------------
def pulisci_importo(val):
    if pd.isna(val):
        return None
    s = str(val).replace("€", "").replace("EUR", "").strip()
    s = re.sub(r"[^\d,.\-]", "", s)
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None

if "importo_stanziato" in df.columns:
    df["importo_stanziato"] = df["importo_stanziato"].apply(pulisci_importo)

# -------------------------------------------------------
# NAVIGAZIONE
# -------------------------------------------------------
lista_pagine = ["Home"] + sorted(df["nome_istituto"].dropna().unique())
st.sidebar.subheader("🧭 Navigazione")
pagina = st.sidebar.radio("Vai a", lista_pagine, key="nav_radio")

# -------------------------------------------------------
# FILTRI GLOBALI
# -------------------------------------------------------
st.sidebar.subheader("🔎 Filtri globali")

filtro_tipologia = st.sidebar.multiselect(
    "Tipologia di intervento",
    sorted(df["tipologia_intervento"].dropna().unique())
)

filtro_manut = st.sidebar.selectbox("Manutenzioni", ["Tutti", "Solo manutenzioni", "Solo altri"])

filtro_comune = st.sidebar.multiselect(
    "Comune", sorted(df["comune"].dropna().unique())
)

df_filt = df.copy()
if filtro_tipologia:
    df_filt = df_filt[df_filt["tipologia_intervento"].isin(filtro_tipologia)]
if filtro_manut == "Solo manutenzioni":
    df_filt = df_filt[df_filt["manut_flag"]]
elif filtro_manut == "Solo altri":
    df_filt = df_filt[~df_filt["manut_flag"]]
if filtro_comune:
    df_filt = df_filt[df_filt["comune"].isin(filtro_comune)]

if df_filt.empty:
    st.warning("Nessun intervento corrisponde ai filtri selezionati.")
    st.stop()

# -------------------------------------------------------
# FORMATTAZIONE IMPORTO
# -------------------------------------------------------
def fmt_eur(x):
    if pd.isna(x):
        return "-"
    return f"€ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# -------------------------------------------------------
# PAGINA HOME
# -------------------------------------------------------
if pagina == "Home":
    st.markdown('<div class="sulcis-card">', unsafe_allow_html=True)
    st.markdown(
        '<div class="sulcis-section-title">🏠 Dashboard generale – Provincia del Sulcis Iglesiente</div>',
        unsafe_allow_html=True,
    )

    # Elenco interventi filtrati
    st.subheader("Elenco interventi (filtrati)")
    colonne_tab = [
        "nome_istituto",
        "codice",
        "comune",
        "tipologia_intervento",
        "manutenzioni",
        "rup",
        "denominazione_intervento",
        "determina",
    ]
    if "importo_stanziato" in df_filt.columns:
        colonne_tab.append("importo_stanziato")

    col_cfg = {}
    if "importo_stanziato" in df_filt.columns:
        col_cfg["importo_stanziato"] = st.column_config.NumberColumn(
            "Importo stanziato", format="€ %,.2f"
        )

    st.dataframe(
        df_filt[colonne_tab],
        use_container_width=True,
        column_config=col_cfg or None,
    )

    # Grafici generali
    col_g1, col_g2 = st.columns(2)
    with col_g1:
        st.subheader("Numero interventi per istituto")
        st.bar_chart(df_filt.groupby("nome_istituto").size())
    with col_g2:
        st.subheader("Manutenzioni vs altri")
        n_m = df_filt[df_filt["manut_flag"]].shape[0]
        n_a = df_filt.shape[0] - n_m
        st.bar_chart(
            pd.DataFrame({"Tipo": ["Manutenzioni", "Altri"], "Valore": [n_m, n_a]}).set_index("Tipo")
        )

    # Riepilogo economico
    st.subheader("💶 Riepilogo economico (importo stanziato)")

    if "importo_stanziato" in df_filt.columns:

        # DF deduplicato per le somme: stessa determina+importo per stesso istituto = 1
        df_rip = df_riepilogo(df_filt)

        col_e1, col_e2 = st.columns(2)

        # Somma per istituto
        with col_e1:
            st.markdown("**Somma importi per istituto**")
            s_ist = (
                df_rip.groupby("nome_istituto")["importo_stanziato"]
                .sum()
                .sort_values(ascending=False)
                .reset_index()
            )
            s_ist["Importo (€)"] = s_ist["importo_stanziato"].map(fmt_eur)
            st.dataframe(
                s_ist[["nome_istituto", "Importo (€)"]],
                use_container_width=True,
            )

        # Somma per tipologia
        with col_e2:
            st.markdown("**Somma importi per tipologia (dedup determina/importo per istituto)**")
            s_tip = (
                df_rip.groupby("tipologia_intervento")["importo_stanziato"]
                .sum()
                .sort_values(ascending=False)
                .reset_index()
            )
            s_tip["Importo (€)"] = s_tip["importo_stanziato"].map(fmt_eur)
            st.dataframe(
                s_tip[["tipologia_intervento", "Importo (€)"]],
                use_container_width=True,
            )

        # ---------------------------------------------------
        # Riepilogo: numero di istituti coinvolti per tipologia
        # ---------------------------------------------------
        st.subheader("🏫 Istituti coinvolti per tipologia")

        istituti_per_tip = (
            df_rip.groupby("tipologia_intervento")["nome_istituto"]
            .nunique()
            .reset_index()
            .rename(columns={"nome_istituto": "Numero istituti"})
            .sort_values("Numero istituti", ascending=False)
        )

        st.dataframe(
            istituti_per_tip,
            use_container_width=True,
        )

        # ---------------------------------------------------
        # DETERMINE ACCORDO/SERVIZIO con scuole coinvolte e quota per scuola
        # ---------------------------------------------------
        st.subheader("📑 Determine Accordo/Servizio (pro-quota per scuola)")

        df_acc = df_filt[df_filt["tipologia_intervento"].str.lower() == "accordo/servizio"].copy()

        tot_quota = 0.0

        if df_acc.empty:
            st.info("Nessuna determina per la tipologia 'Accordo/Servizio' nei filtri correnti.")
        else:
            df_acc["determina_norm"] = df_acc["determina"].astype(str).str.strip().str.lower()

            # dedup per istituto/determina/importo
            df_acc_uni = df_acc.drop_duplicates(
                subset=["codice", "determina_norm", "importo_stanziato"]
            )

            # aggrego per determina: somma importi + numero scuole distinte
            det_acc = (
                df_acc_uni.groupby(["determina_norm", "determina"])
                .agg(
                    importo_stanziato=("importo_stanziato", "sum"),
                    numero_scuole=("codice", "nunique"),
                )
                .reset_index()
                .sort_values("importo_stanziato", ascending=False)
            )

            # quota per scuola = importo totale / numero_scuole
            det_acc["importo_per_scuola"] = det_acc["importo_stanziato"] / det_acc["numero_scuole"]

            det_acc["Importo totale (€)"] = det_acc["importo_stanziato"].map(fmt_eur)
            det_acc["Quota per scuola (€)"] = det_acc["importo_per_scuola"].map(fmt_eur)

            st.dataframe(
                det_acc[["determina", "numero_scuole", "Importo totale (€)", "Quota per scuola (€)"]],
                use_container_width=True,
            )

            # Totale complessivo delle quote per scuola
            tot_quota = det_acc["importo_per_scuola"].sum()
            st.success(f"Totale complessivo quote Accordo/Servizio (somma quote per scuola): {fmt_eur(tot_quota)}")

        # ---------------------------------------------------
        # Somma per manutenzione (VERO / FALSO) usando tot_quota per manutenzioni
        # ---------------------------------------------------
        st.markdown("**Somma importi per manutenzione (VERO / FALSO)**")

        s_man = (
            df_rip.groupby("manut_flag")["importo_stanziato"]
            .sum()
            .reset_index()
        )
        s_man["manutenzione"] = s_man["manut_flag"].map(
            {True: "VERO (manutenzioni)", False: "FALSO (altri interventi)"}
        )

        mask_vero = s_man["manutenzione"] == "VERO (manutenzioni)"
        if mask_vero.any():
            s_man.loc[mask_vero, "importo_stanziato"] = tot_quota

        s_man["Importo (€)"] = s_man["importo_stanziato"].map(fmt_eur)
        st.dataframe(
            s_man[["manutenzione", "Importo (€)"]],
            use_container_width=True,
        )

        totale_generale = s_man["importo_stanziato"].sum()
        st.success(f"**Totale generale stanziato (con quote manutenzioni): {fmt_eur(totale_generale)}**")

    else:
        st.info("Colonna 'importo stanziato' non presente.")

    st.markdown('</div>', unsafe_allow_html=True)

# -------------------------------------------------------
# PAGINE ISTITUTO
# -------------------------------------------------------
else:
    istituto_sel = pagina
    df_ist = df_filt[df_filt["nome_istituto"] == istituto_sel]

    st.markdown('<div class="sulcis-card">', unsafe_allow_html=True)
    st.markdown(
        f'<div class="sulcis-section-title">🏫 {istituto_sel} – Provincia del Sulcis Iglesiente</div>',
        unsafe_allow_html=True,
    )

    row_ist = istituti[istituti["nome_istituto"] == istituto_sel].head(1)
    if not row_ist.empty:
        c1, c2 = st.columns(2)
        with c1:
            st.markdown(f"**Comune:** {row_ist.iloc[0].get('comune', '')}")
        with c2:
            st.markdown(f"**Indirizzo:** {row_ist.iloc[0].get('indirizzo', '')}")

    colonne_base = [
        "tipologia_intervento",
        "manutenzioni",
        "rup",
        "denominazione_intervento",
        "determina",
    ]
    if "importo_stanziato" in df_ist.columns:
        colonne_base.append("importo_stanziato")

    col_cfg_ist = {}
    if "importo_stanziato" in df_ist.columns:
        col_cfg_ist["importo_stanziato"] = st.column_config.NumberColumn(
            "Importo stanziato", format="€ %,.2f"
        )

    # 1) tutti
    st.subheader("📋 Interventi (tutti)")
    st.dataframe(
        df_ist[colonne_base],
        use_container_width=True,
        column_config=col_cfg_ist or None,
    )

    # 2) manutenzioni
    st.subheader("🛠️ Interventi di manutenzione (VERO)")
    df_m = df_ist[df_ist["manut_flag"]]
    if df_m.empty:
        st.info("Nessuna manutenzione per questo istituto.")
    else:
        st.dataframe(
            df_m[colonne_base],
            use_container_width=True,
            column_config=col_cfg_ist or None,
        )

    # 3) non manutenzioni
    st.subheader("📋 Interventi diversi dalle manutenzioni (FALSO)")
    df_nm = df_ist[~df_ist["manut_flag"]]
    if df_nm.empty:
        st.info("Nessun intervento non di manutenzione per questo istituto.")
    else:
        st.dataframe(
            df_nm[colonne_base],
            use_container_width=True,
            column_config=col_cfg_ist or None,
        )

    # Grafici istituto
    st.subheader("📊 Grafici istituto")
    cg1, cg2 = st.columns(2)
    with cg1:
        st.markdown("**Interventi per tipologia**")
        st.bar_chart(df_ist.groupby("tipologia_intervento").size())
    with cg2:
        st.markdown("**Manutenzioni vs altri**")
        n_mi = df_ist[df_ist["manut_flag"]].shape[0]
        n_ai = df_ist.shape[0] - n_mi
        st.bar_chart(
            pd.DataFrame({"Tipo": ["Manutenzioni", "Altri"], "Valore": [n_mi, n_ai]}).set_index("Tipo")
        )

    # ---------------------------------------------------
    # Riepilogo per tipologia (istituto corrente)
    # ---------------------------------------------------
    st.subheader("📌 Riepilogo interventi per tipologia (istituto)")

    riepilogo_tip_ist = (
        df_ist.groupby("tipologia_intervento")
        .size()
        .reset_index(name="Numero interventi")
        .sort_values("Numero interventi", ascending=False)
    )

    st.dataframe(
        riepilogo_tip_ist,
        use_container_width=True,
    )

    # ---------------------------------------------------
    # PDF ISTITUTO (OPZIONALE, con logica quote Accordo/Servizio)
    # ---------------------------------------------------
    if REPORTLAB_AVAILABLE:
        def crea_pdf(data: pd.DataFrame, nome: str) -> BytesIO:
            buffer = BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=A4)
            styles = getSampleStyleSheet()
            elements = []

            # Logo
            try:
                resp = requests.get(LOGO_URL, timeout=5)
                if resp.status_code == 200:
                    logo = RLImage(BytesIO(resp.content), width=40, height=40)
                    elements.append(logo)
                    elements.append(Spacer(1, 6))
            except Exception:
                pass

            elements.append(Paragraph(f"Report Istituto: {nome}", styles["Title"]))
            elements.append(Paragraph("Provincia del Sulcis Iglesiente", styles["Normal"]))
            elements.append(Spacer(1, 12))

            # Dedup locale per questo istituto
            data_local = data.copy()
            if "importo_stanziato" in data_local.columns:
                data_local["determina_norm"] = data_local["determina"].astype(str).str.strip().str.lower()
                data_rip_ist = df_riepilogo(data_local)
            else:
                data_rip_ist = data_local

            # Quote Accordo/Servizio per questo istituto (coerenti con la Home)
            quota_acc_manut = 0.0
            if "importo_stanziato" in data_rip_ist.columns:
                df_acc_all = df_filt[df_filt["tipologia_intervento"].str.lower() == "accordo/servizio"].copy()
                if not df_acc_all.empty:
                    df_acc_all["determina_norm"] = df_acc_all["determina"].astype(str).str.strip().str.lower()
                    df_acc_uni_all = df_acc_all.drop_duplicates(
                        subset=["codice", "determina_norm", "importo_stanziato"]
                    )
                    det_acc_all = (
                        df_acc_uni_all.groupby(["determina_norm", "determina"])
                        .agg(
                            importo_stanziato=("importo_stanziato", "sum"),
                            numero_scuole=("codice", "nunique"),
                        )
                        .reset_index()
                    )
                    det_acc_all["importo_per_scuola"] = (
                        det_acc_all["importo_stanziato"] / det_acc_all["numero_scuole"]
                    )

                    det_acc_merge = df_acc_all.merge(
                        det_acc_all[["determina_norm", "importo_per_scuola"]],
                        on="determina_norm",
                        how="left",
                        suffixes=("", "_quota"),
                    )

                    df_acc_ist = det_acc_merge[det_acc_merge["nome_istituto"] == nome]
                    quota_acc_manut = df_acc_ist["importo_per_scuola"].sum()

            # Riepilogo numerico di base
            n_tot = len(data_local)
            n_mn = data_local[data_local["manut_flag"]].shape[0]
            elements.append(
                Paragraph(
                    f"Interventi totali: {n_tot} – Manutenzioni: {n_mn} – Altri: {n_tot - n_mn}",
                    styles["Normal"],
                )
            )
            elements.append(Spacer(1, 12))

            # Riepilogo economico istituto con stessa logica della Home
            if "importo_stanziato" in data_rip_ist.columns:
                # Somma NON manutenzioni (sempre raw dedup)
                s_al_raw = data_rip_ist[~data_rip_ist["manut_flag"]]["importo_stanziato"].sum()

                # Somma manutenzioni: sostituisco con quota_acc_manut
                s_mn_eff = quota_acc_manut

                s_tot_eff = s_mn_eff + s_al_raw

                txt = (
                    f"Importo stanziato totale (con quote Accordo/Servizio): {fmt_eur(s_tot_eff)} "
                    f"(Manutenzioni: {fmt_eur(s_mn_eff)} – Altri: {fmt_eur(s_al_raw)})"
                )
                elements.append(Paragraph(txt, styles["Normal"]))
                elements.append(Spacer(1, 12))

            # Tabella dettagli interventi (importi raw, come da CSV/filtri)
            hs = styles["Heading5"]
            cs = styles["Normal"]
            cs.fontSize = 8

            table_data = [[
                Paragraph("Tipologia", hs),
                Paragraph("Manut.", hs),
                Paragraph("RUP", hs),
                Paragraph("Intervento", hs),
                Paragraph("Determina", hs),
                Paragraph("Importo", hs),
            ]]

            for _, row in data_local.iterrows():
                if "importo_stanziato" in row and pd.notna(row["importo_stanziato"]):
                    imp_txt = fmt_eur(row["importo_stanziato"])
                else:
                    imp_txt = "-"
                table_data.append([
                    Paragraph(str(row["tipologia_intervento"]), cs),
                    Paragraph("Sì" if row["manut_flag"] else "No", cs),
                    Paragraph(str(row.get("rup", "")), cs),
                    Paragraph(str(row["denominazione_intervento"]), cs),
                    Paragraph(str(row["determina"]), cs),
                    Paragraph(imp_txt, cs),
                ])

            t = Table(table_data, repeatRows=1, colWidths=[65, 30, 60, 200, 90, 80])
            t.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, 0), 9),
                ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                ("ALIGN", (1, 1), (1, -1), "CENTER"),
                ("ALIGN", (5, 1), (5, -1), "RIGHT"),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 3),
                ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                ("TOPPADDING", (0, 0), (-1, -1), 2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ]))
            elements.append(t)

            doc.build(elements)
            buffer.seek(0)
            return buffer

        pdf = crea_pdf(df_ist, istituto_sel)
        st.download_button(
            label="📄 Scarica report PDF istituto",
            data=pdf,
            file_name=f"report_{istituto_sel}.pdf",
            mime="application/pdf",
        )
    else:
        st.info("Generazione PDF disabilitata: modulo 'reportlab' non disponibile.")

    st.markdown('</div>', unsafe_allow_html=True)
