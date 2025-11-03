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
    Aggiunge doppi apici (""val"") solo se c'è un valore non vuoto/non 'nan'.
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

def pick_ci(df: pd.DataFrame, candidates) -> pd.Series | None:
    """
    Ritorna la prima colonna trovata tra i nomi candidati (case-insensitive), altrimenti None.
    """
    cols = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        key = str(cand).strip().lower()
        if key in cols:
            return df[cols[key]]
    return None
def extract_sam_from_description(desc_val: str) -> str:
    """
    Estrae il samaccountname dalla colonna Description con formato:
    "... - QUALCOSA - {sam} - DATA ORA"
    Strategia: split per ' - ' e prendi il token immediatamente prima dell'ultimo.
    Rimuove eventuali parentesi graffe { } o < >.
    """
    s = normalize_str(desc_val)
    if not s:
        return ""
    parts = [p.strip() for p in s.split(" - ")]
    if len(parts) >= 2:
        candidate = parts[-2]
        if (candidate.startswith("{") and candidate.endswith("}")) or \
           (candidate.startswith("<") and candidate.endswith(">")):
            candidate = candidate[1:-1].strip()
        return candidate
    return ""

# -----------------------------
# Input: file & nomi output
# -----------------------------
st.subheader("1) Carica il file Excel estr_device (.xlsx)")
uploaded_device = st.file_uploader("Seleziona il file estr_device", type=["xlsx"], key="estr_device")

st.subheader("1b) (Opzionale) Carica il file Excel estr_dati (.xlsx) per MAIL/MOBILE/DISPLAYNAME")
uploaded_dati = st.file_uploader("Seleziona il file estr_dati", type=["xlsx"], key="estr_dati")

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
    if uploaded_device is None:
        st.error("Per favore carica prima il file Excel estr_device.")
        st.stop()

    # ============== Lettura estr_device ==============
    try:
        raw_dev = pd.read_excel(uploaded_device, dtype=str)
    except Exception as e:
        st.error(f"Errore nel leggere estr_device: {e}")
        st.stop()

    raw_dev.columns = [str(c).strip() for c in raw_dev.columns]

    # Colonne necessarie in estr_device
    try:
        dev_description   = get_col_case_insensitive(raw_dev, "Description")          # J -> estrae sam
        dev_mail          = get_col_case_insensitive(raw_dev, "Mail")                 # K -> fallback add_mail
        dev_mobile        = get_col_case_insensitive(raw_dev, "Mobile")               # L -> fallback add_mobile
        dev_upn           = get_col_case_insensitive(raw_dev, "userPrincipalName")    # N -> fallback add_userprincipalname
        dev_name_oldpc    = get_col_case_insensitive(raw_dev, "Name")                 # vecchio asset
        dev_enabled       = get_col_case_insensitive(raw_dev, "Enabled")              # F
        dev_dn            = get_col_case_insensitive(raw_dev, "DistinguishedName")    # E
    except KeyError as ke:
        st.error(str(ke))
        st.stop()

    # Filtro Enabled = True
    dev_mask = dev_enabled.astype(str).str.strip().str.lower() == "true"
    fdev = raw_dev.loc[dev_mask].copy()

    if fdev.empty:
        st.warning("Nessun record 'Enabled = True' trovato in estr_device: nulla da elaborare.")
        st.stop()

    # Rilettura dal df filtrato
    f_description   = get_col_case_insensitive(fdev, "Description").astype(str)
    f_mail          = get_col_case_insensitive(fdev, "Mail").astype(str)
    f_mobile        = get_col_case_insensitive(fdev, "Mobile").astype(str)
    f_upn           = get_col_case_insensitive(fdev, "userPrincipalName").astype(str)
    f_name_oldpc    = get_col_case_insensitive(fdev, "Name").astype(str)
    f_dn            = get_col_case_insensitive(fdev, "DistinguishedName").astype(str)

    estr_df = pd.DataFrame({
        "samaccountname": f_description.map(extract_sam_from_description),
        "mail_dev":       f_mail.map(normalize_str),         # fallback
        "mobile_dev":     f_mobile.map(normalize_str),       # fallback
        "upn_dev":        f_upn.map(normalize_str),          # fallback per add_userprincipalname
        "old_computer":   f_name_oldpc.map(normalize_str),   # vecchio asset
        "dn":             f_dn.map(normalize_str)
    })
    estr_df["sam_norm"] = estr_df["samaccountname"].map(lower_norm)

    # ============== Lettura (opzionale) estr_dati ==============
    # Obiettivo: ricavare mail/mobile/displayname più "autoritative" e fare merge su sam_norm
    dati_loaded = False
    if uploaded_dati is not None:
        try:
            raw_dati = pd.read_excel(uploaded_dati, dtype=str)
        except Exception as e:
            st.error(f"Errore nel leggere estr_dati: {e}")
            st.stop()

        raw_dati.columns = [str(c).strip() for c in raw_dati.columns]

        # Candidati come da storico:
        # - SAM: "SamAccountName"/"sAMAccountName" (A)
        # - Mobile: "Mobile" (E)
        # - mail: "mail"/"Mail"/"e-mail"/"email" (J)
        # - DisplayName: "DisplayName" (preferito), fallback "Name"/"cn"
        sam_dati        = pick_ci(raw_dati, ["SamAccountName", "sAMAccountName"])
        mail_dati       = pick_ci(raw_dati, ["mail", "Mail", "e-mail", "email"])
        mobile_dati     = pick_ci(raw_dati, ["Mobile", "mobile"])
        display_dati    = pick_ci(raw_dati, ["DisplayName", "displayName", "Display Name", "Name", "cn"])

        if sam_dati is None:
            st.error("In estr_dati non è stata trovata la colonna 'SamAccountName'/'sAMAccountName'. Impossibile allineare i dati.")
            st.stop()

        # Normalizza serie (le altre possono essere None -> serie vuote)
        sam_dati = sam_dati.astype(str)

        if mail_dati is None:
            mail_dati = pd.Series([""] * len(raw_dati))
            st.warning("In estr_dati non è stata trovata la colonna 'mail'/'Mail'/'e-mail'/'email'. Userò il fallback da estr_device.")
        else:
            mail_dati = mail_dati.astype(str)

        if mobile_dati is None:
            mobile_dati = pd.Series([""] * len(raw_dati))
            st.warning("In estr_dati non è stata trovata la colonna 'Mobile'. Userò il fallback da estr_device.")
        else:
            mobile_dati = mobile_dati.astype(str)

        if display_dati is None:
            display_dati = pd.Series([""] * len(raw_dati))
            st.warning("In estr_dati non è stata trovata la colonna 'DisplayName' (verrà usato userPrincipalName da estr_device come fallback).")
        else:
            display_dati = display_dati.astype(str)

        dati_map = pd.DataFrame({
            "sam_norm":        sam_dati.map(lower_norm),
            "mail_dati":       mail_dati.map(normalize_str),
            "mobile_dati":     mobile_dati.map(normalize_str),
            "displayname_dati":display_dati.map(normalize_str),
        })

        # Deduplica per sam_norm mantenendo la prima occorrenza
        dati_map = dati_map.drop_duplicates(subset=["sam_norm"], keep="first")

        # Merge su estr_df
        estr_df = estr_df.merge(dati_map, on="sam_norm", how="left")
        dati_loaded = True
    else:
        # Se non caricato estr_dati, crea colonne vuote per uniformità
        estr_df["mail_dati"] = ""
        estr_df["mobile_dati"] = ""
        estr_df["displayname_dati"] = ""

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
            add_mail_val = ""
            add_mobile_val = ""
            add_display_val = ""  # per add_userprincipalname
            old_pc = ""
            dn_val = ""
        else:
            rec = match.iloc[0]

            # Precedenza: estr_dati -> estr_device
            mail_pref       = normalize_str(rec.get("mail_dati", "")) or normalize_str(rec.get("mail_dev", ""))
            mobile_pref     = normalize_str(rec.get("mobile_dati", "")) or normalize_str(rec.get("mobile_dev", ""))
            display_pref    = normalize_str(rec.get("displayname_dati", "")) or normalize_str(rec.get("upn_dev", ""))

            add_mail_val    = mail_pref                                  # non quotato
            add_mobile_val  = quote_if_value(mobile_pref)                 # quotato se presente
            add_display_val = quote_if_value(display_pref)                # add_userprincipalname (da DisplayName o UPN)
            old_pc          = normalize_str(rec.get("old_computer", ""))
            dn_val          = normalize_str(rec.get("dn", ""))

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

        # 2) AGGIUNTA: usa i dati (preferendo estr_dati)
        row_add = [""] * 10
        row_add[0] = nuovo_pc
        row_add[2] = add_mail_val
        row_add[4] = add_mobile_val
        row_add[6] = add_display_val  # <- add_userprincipalname = DisplayName (estr_dati) -> fallback UPN (estr_device)
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

    if not dati_loaded:
        st.info("Non hai caricato 'estr_dati': userò i valori di 'estr_device' per mail/mobile e userPrincipalName per add_userprincipalname (fallback).")

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
