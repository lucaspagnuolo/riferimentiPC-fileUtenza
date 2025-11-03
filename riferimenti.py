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

def get_col_case_insensitive(df: pd.DataFrame, wanted: str) -> pd.Series:
    """
    Restituisce la colonna per nome (case-insensitive) o solleva un errore se non presente.
    """
    wanted_l = wanted.strip().lower()
    for c in df.columns:
        if str(c).strip().lower() == wanted_l:
            return df[c]
    raise KeyError(f"Colonna richiesta '{wanted}' non trovata nell'Excel caricato.")

def extract_sam_from_mail(mail_val: str) -> str:
    """
    Estrae il samaccountname da un indirizzo email del tipo {sam}@consip.it.
    """
    s = normalize_str(mail_val)
    if "@" in s:
        return s.split("@", 1)[0].strip()
    return s

# -----------------------------
# Input: file & nomi output
# -----------------------------
st.subheader("1) Carica il file Excel estr_device (.xlsx)")
uploaded = st.file_uploader("Seleziona il file estr_device", type=["xlsx"])

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
        st.error("Per favore carica prima il file Excel estr_device.")
        st.stop()

    # Leggi Excel
    try:
        raw_df = pd.read_excel(uploaded, dtype=str)
    except Exception as e:
        st.error(f"Errore nel leggere l'Excel: {e}")
        st.stop()

    # Normalizza intestazioni
    raw_df.columns = [str(c).strip() for c in raw_df.columns]

    # ============================
    # Lettura colonne richieste
    # ============================
    try:
        col_mail            = get_col_case_insensitive(raw_df, "Mail")
        col_upn             = get_col_case_insensitive(raw_df, "userPrincipalName")
        col_mobile          = get_col_case_insensitive(raw_df, "Mobile")
        col_name_oldpc      = get_col_case_insensitive(raw_df, "Name")  # vecchio asset
        col_enabled         = get_col_case_insensitive(raw_df, "Enabled")
        col_distinguished   = get_col_case_insensitive(raw_df, "DistinguishedName")
    except KeyError as ke:
        st.error(str(ke))
        st.stop()

    # ============================
    # Filtro: Enabled == True
    # ============================
    enabled_mask = col_enabled.astype(str).str.strip().str.lower() == "true"
    filtered_df = raw_df.loc[enabled_mask].copy()

    if filtered_df.empty:
        st.warning("Nessun record 'Enabled = True' trovato in estr_device: nulla da elaborare.")
        st.stop()

    # Rilettura colonne dal df filtrato (per avere gli indici coerenti)
    f_mail          = get_col_case_insensitive(filtered_df, "Mail").astype(str)
    f_upn           = get_col_case_insensitive(filtered_df, "userPrincipalName").astype(str)
    f_mobile        = get_col_case_insensitive(filtered_df, "Mobile").astype(str)
    f_name_oldpc    = get_col_case_insensitive(filtered_df, "Name").astype(str)
    f_dn            = get_col_case_insensitive(filtered_df, "DistinguishedName").astype(str)

    # Costruzione DF di lavoro
    estr_df = pd.DataFrame({
        "samaccountname": f_mail.map(extract_sam_from_mail),   # ricavato da Mail
        "mail":           f_mail.map(normalize_str),           # mail originale
        "mobile":         f_mobile.map(normalize_str),
        "display":        f_upn.map(normalize_str),            # Name (display) = userPrincipalName
        "old_computer":   f_name_oldpc.map(normalize_str),     # vecchio asset
        "dn":             f_dn.map(normalize_str)
    })

    estr_df["sam_norm"] = estr_df["samaccountname"].map(lower_norm)

    # -----------------------------
    # Header dei CSV
    # -----------------------------
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
    alerts = []
    valid_pairs = 0

    for _, r in pairs_df.iterrows():
        nuovo_pc = normalize_str(r.get("NuovoPC", ""))
        utenza   = normalize_str(r.get("samaccountname", ""))

        if not nuovo_pc or not utenza:
            continue
        valid_pairs += 1

        # Match su samaccountname (case-insensitive)
        match = estr_df[estr_df["sam_norm"] == utenza.lower()]

        if match.empty:
            warnings.append(f"• Utente '{utenza}' non trovato tra i record con Enabled = True in estr_device.")
            mail = ""
            mobile_q  = ""
            display_q = ""
            old_pc    = ""
            dn_val    = ""
        else:
            rec = match.iloc[0]
            mail      = normalize_str(rec.get("mail", ""))
            mobile_q  = quote_if_value(normalize_str(rec.get("mobile", "")))
            display_q = quote_if_value(normalize_str(rec.get("display", "")))
            old_pc    = normalize_str(rec.get("old_computer", ""))
            dn_val    = normalize_str(rec.get("dn", ""))

            # Alert: vecchio asset con OU=PDL in dismissione nel DN
            if old_pc and "ou=pdl in dismissione" in dn_val.lower():
                alerts.append(
                    f"⚠️ Vecchio asset '{old_pc}' è in 'OU=PDL in dismissione' (DN: {dn_val}) per l'utenza '{utenza}'."
                )

        # -----------------------------
        # CSV 1: RIFERIMENTI
        # -----------------------------

        # 1) RIMOZIONE: SOLO 'SI' + Computer=old_pc (se presente)
        if old_pc:
            row_remove = [""] * 10
            row_remove[0] = old_pc   # Computer
            row_remove[3] = "SI"     # remove_mail
            row_remove[5] = "SI"     # remove_mobile
            row_remove[7] = "SI"     # remove_userprincipalname
            rows_rif.append(row_remove)

        # 2) AGGIUNTA: usa i dati letti da estr_device
        row_add = [""] * 10
        row_add[0] = nuovo_pc      # Computer
        row_add[2] = mail          # add_mail (non quotato)
        row_add[4] = mobile_q      # add_mobile (quotato se presente)
        row_add[6] = display_q     # add_userprincipalname (quotato se presente)
        rows_rif.append(row_add)

        # -----------------------------
        # CSV 2: DESCRITION (23 colonne)
        # -----------------------------
        row_desc = [""] * 23
        row_desc[0]  = utenza      # sAMAccountName
        row_desc[11] = nuovo_pc    # Description
        rows_desc.append(row_desc)

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

    # Esito & Messaggi
    st.success(f"CSV generati: {file1_name} (righe: {len(rows_rif)}) e {file2_name} (righe: {len(rows_desc)})")

    if warnings:
        st.warning("**Avvisi (match non trovati):**\n" + "\n".join(warnings))

    if alerts:
        st.info("**Alert (vecchi asset in OU=PDL in dismissione):**\n" + "\n".join(alerts))

    # Download + Anteprime
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
``
