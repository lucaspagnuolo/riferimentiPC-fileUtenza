# app.py
import streamlit as st
import pandas as pd
from io import StringIO
from datetime import date
import csv

# -----------------------------
# Configurazione pagina
# -----------------------------
st.set_page_config(page_title="Generatore CSV Riferimenti & Descrizione", layout="wide")

st.title("Generatore CSV per Asset PC (Riferimenti + Descrition)")
st.caption("Unisce e generalizza i tuoi codice1/codice2, con supporto a più combinazioni NuovoPC–samaccountname.")

# -----------------------------
# Utility
# -----------------------------
def today_yyyymmdd() -> str:
    return date.today().strftime("%Y%m%d")

def normalize_str(val) -> str:
    return "" if pd.isna(val) else str(val).strip()

def lower_norm(val) -> str:
    return normalize_str(val).lower()

def quote_if_value(val: str) -> str:
    """
    Aggiunge doppi apici come nei tuoi script originali
    (vuoto o 'nan' -> stringa vuota).
    """
    if val is None:
        return ""
    sval = str(val).strip()
    if sval == "" or sval.lower() == "nan":
        return ""
    return f'"{sval}"'

def pick_column(series_names, df: pd.DataFrame, fallback_idx=None) -> pd.Series:
    """
    Prova a trovare la colonna per nome (case-insensitive).
    Se non la trova, usa il fallback per indice (come da posizioni A,B,E,J,Q).
    """
    cols_lower = {str(c).strip().lower(): c for c in df.columns}
    for candidate in series_names:
        key = str(candidate).strip().lower()
        if key in cols_lower:
            return df[cols_lower[key]]
    if fallback_idx is not None and 0 <= fallback_idx < df.shape[1]:
        return df.iloc[:, fallback_idx]
    # Nessun match: torna una serie vuota così non rompe il flusso
    return pd.Series([None] * len(df))

# -----------------------------
# Input: file & nomi output
# -----------------------------
st.subheader("1) Carica il file Excel estr_dati (.xlsx)")
uploaded = st.file_uploader("Seleziona il file estr_dati", type=["xlsx"])

st.subheader("2) Nomi file di output (puoi cambiarli)")
default_file1 = f"{today_yyyymmdd()}_Computer_riferimenti.csv"
default_file2 = f"{today_yyyymmdd()}_Descrition.csv"
file1_name = st.text_input("Nome file CSV Riferimenti", value=default_file1)
file2_name = st.text_input("Nome file CSV Descrition", value=default_file2)

# -----------------------------
# Editor combinazioni NuovoPC–samaccountname
# -----------------------------
st.subheader("3) Combinazioni da elaborare")
st.caption("Aggiungi una riga per ogni coppia: **Nome asset del nuovo PC** e **samaccountname**.")

if "pairs_df" not in st.session_state:
    st.session_state["pairs_df"] = pd.DataFrame([{"NuovoPC": "", "samaccountname": ""}])

pairs_df = st.data_editor(
    st.session_state["pairs_df"],
    num_rows="dynamic",
    use_container_width=True,
    key="pairs_editor",
    column_config={
        "NuovoPC": st.column_config.TextColumn(label="Nome asset del nuovo PC"),
        "samaccountname": st.column_config.TextColumn(label="samaccountname (utenza)"),
    }
)
st.session_state["pairs_df"] = pairs_df

# -----------------------------
# Bottone di generazione
# -----------------------------
generate = st.button("Genera CSV")

if generate:
    if uploaded is None:
        st.error("Per favore carica prima il file Excel estr_dati.")
        st.stop()

    # Leggi Excel
    try:
        raw_df = pd.read_excel(uploaded, dtype=str)
    except Exception as e:
        st.error(f"Errore nel leggere l'Excel: {e}")
        st.stop()

    # Normalizza header
    raw_df.columns = [str(c).strip() for c in raw_df.columns]

    # Mappatura colonne con fallback a posizioni:
    # A: SamAccountName (0), B: Name (1), E: Mobile (4), J: mail (9), Q: Description (16)
    col_sam = pick_column(["SamAccountName", "sAMAccountName"], raw_df, fallback_idx=0)
    col_name = pick_column(["Name"], raw_df, fallback_idx=1)
    col_mobile = pick_column(["Mobile"], raw_df, fallback_idx=4)
    col_mail = pick_column(["mail", "Mail", "e-mail", "email"], raw_df, fallback_idx=9)
    col_desc_old = pick_column(["Description", "Descrizione"], raw_df, fallback_idx=16)

    estr_df = pd.DataFrame({
        "samaccountname": col_sam.astype(str),
        "name": col_name.astype(str),
        "mobile": col_mobile.astype(str),
        "mail": col_mail.astype(str),
        "description_old": col_desc_old.astype(str),
    })

    # Colonna normalizzata per match case-insensitive
    estr_df["sam_norm"] = estr_df["samaccountname"].map(lower_norm)

    # Header come nei tuoi script
    header_rif = [
        "Computer", "OU",
        "add_mail", "remove_mail",
        "add_mobile", "remove_mobile",
        "add_userprincipalname", "remove_userprincipalname",
        "disable", "moveToOU"
    ]

    header_desc = [
        "sAMAccountName", "Creation", "OU", "Name", "DisplayName", "cn", "GivenName", "Surname",
        "employeeNumber", "employeeID", "department", "Description", "passwordNeverExpired",
        "ExpireDate", "userprincipalname", "mail", "mobile", "RimozioneGruppo", "InserimentoGruppo",
        "disable", "moveToOU", "telephoneNumber", "company"
    ]

    rows_rif = []
    rows_desc = []
    warnings = []

    # Cicla le coppie
    valid_pairs = 0
    for _, r in pairs_df.iterrows():
        nuovo_pc = normalize_str(r.get("NuovoPC", ""))
        utenza = normalize_str(r.get("samaccountname", ""))

        if not nuovo_pc or not utenza:
            # salta righe vuote
            continue

        valid_pairs += 1

        # Cerca match SamAccountName
        match = estr_df[estr_df["sam_norm"] == utenza.lower()]
        if match.empty:
            warnings.append(f"• Utente '{utenza}' non trovato nel file estr_dati: campi mail/mobile/Name lasciati vuoti.")
            mail = ""  # come nel tuo codice1: mail non quotata
            mobile_q = quote_if_value("")  # mobile quotato se presente
            display_q = quote_if_value("")  # Name quotato se presente
            old_pc = ""
        else:
            rec = match.iloc[0]
            mail = normalize_str(rec["mail"])  # non quotato, come nel codice1
            mobile_q = quote_if_value(normalize_str(rec["mobile"]))  # quotato nel codice1
            display_q = quote_if_value(normalize_str(rec["name"]))   # quotato nel codice1
            old_pc = normalize_str(rec["description_old"])
            if old_pc.lower() == "nan":
                old_pc = ""

        # CSV 1: riga di aggiunta (sempre)
        rows_rif.append([
            nuovo_pc, "",           # Computer, OU
            mail, "",               # add_mail, remove_mail
            mobile_q, "",           # add_mobile, remove_mobile
            display_q, "",          # add_userprincipalname, remove_userprincipalname
            "", ""                  # disable, moveToOU
        ])

        # CSV 1: riga di rimozione se c'è il vecchio asset (Description)
        if old_pc:
            rows_rif.append([
                old_pc, "",         # Computer, OU
                "", mail,           # add_mail, remove_mail
                "", mobile_q,       # add_mobile, remove_mobile
                "", display_q,      # add_userprincipalname, remove_userprincipalname
                "", ""              # disable, moveToOU
            ])

        # CSV 2: come nel tuo codice2 -> solo sAMAccountName e Description (con doppi apici)
        rows_desc.append([
            quote_if_value(utenza),     # sAMAccountName
            "", "", "", "", "", "", "", "", "", "",
            quote_if_value(nuovo_pc),   # Description
            "", "", "", "", "", "", "", "", "", ""
        ])

    if valid_pairs == 0:
        st.warning("Nessuna combinazione valida: inserisci almeno una riga con **NuovoPC** e **samaccountname**.")
        st.stop()

    # Serializza CSV in memoria
    buf1 = StringIO()
    w1 = csv.writer(buf1, lineterminator="\n")
    w1.writerow(header_rif)
    w1.writerows(rows_rif)

    buf2 = StringIO()
    w2 = csv.writer(buf2, lineterminator="\n")
    w2.writerow(header_desc)
    w2.writerows(rows_desc)

    # Esito & Download
    st.success(f"CSV generati: {file1_name} (righe: {len(rows_rif)}) e {file2_name} (righe: {len(rows_desc)})")

    if warnings:
        st.warning("\n".join(warnings))

    col_a, col_b = st.columns(2)
    with col_a:
        st.download_button(
            "⬇️ Scarica CSV Riferimenti",
            data=buf1.getvalue().encode("utf-8"),
            file_name=file1_name,
            mime="text/csv"
        )
        st.markdown("**Anteprima Riferimenti**")
        st.dataframe(pd.DataFrame(rows_rif, columns=header_rif).head(50), use_container_width=True)

    with col_b:
        st.download_button(
            "⬇️ Scarica CSV Descrition",
            data=buf2.getvalue().encode("utf-8"),
            file_name=file2_name,
            mime="text/csv"
        )
        st.markdown("**Anteprima Descrition**")
        st.dataframe(pd.DataFrame(rows_desc, columns=header_desc).head(50), use_container_width=True)
