import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Tariffario Noli Marittimi", layout="wide", page_icon="🚢")

DB_FILE = "database_noli_centralizzato.csv"

COMPAGNIE_SUPPORTATE = [
    "MSC", "CMA", "HAPAG", "MAERSK", "EVERGREEN", "MESSINA", "GRIMALDI", 
    "COSCO", "YML", "ONE", "AKKON", "ARKAS", "TARROS", "HMM", "OOCL", 
    "SAHEL", "SCI LINE", "ZIM", "COSIARMA", "MARFRET", "BORCHARD", 
    "COTUNAV", "MAGUISA"
]

TRADE_SUPPORTATI = [
    "IPBC", "RED SEA", "EAF", "MIDDLE EAST", "FAR EAST", "MEDITERRANEAN", "GENERIC"
]

# Elenco di riferimento per identificare le righe dei porti italiani nei listini verticali
CITTA_ITALIANE_RILEVAMENTO = [
    "GENOVA", "LA SPEZIA", "CIVITAVECCHIA", "NAPOLI", "SALERNO", 
    "GIOIA TAURO", "CAGLIARI", "BARI", "ANCONA", "RAVENNA", "VENEZIA", "TRIESTE"
]

def normalizza_stringa_porto(testo_grezzo):
    """Pulisce la stringa rimuovendo spazi doppi e ritorni a capo"""
    t = str(testo_grezzo).strip().replace("\n", " ")
    return " ".join(t.split()).upper()

def carica_database():
    if os.path.exists(DB_FILE):
        df = pd.read_csv(DB_FILE)
        for col_testo in ["POL", "POD", "Compagnia", "Trade", "Container", "Descrizione_Addizionali", "Descrizione_Spese_Imbarco", "Free_Time", "Validità", "Note", "Origine", "Valuta_Nolo", "Valuta_Addizionali", "Valuta_Spese_Imbarco"]:
            if col_testo in df.columns:
                df[col_testo] = df[col_testo].astype(str).replace("nan", "")
        for col_num in ["Nolo", "Addizionali", "Totale_Nolo", "Spese_Imbarco", "BL"]:
            if col_num in df.columns:
                df[col_num] = pd.to_numeric(df[col_num], errors='coerce').fillna(0.0)
        return df
    else:
        return pd.DataFrame(columns=[
            "POL", "POD", "Compagnia", "Trade", "Container", 
            "Nolo", "Valuta_Nolo", 
            "Addizionali", "Valuta_Addizionali", "Descrizione_Addizionali", 
            "Totale_Nolo", 
            "Spese_Imbarco", "Valuta_Spese_Imbarco", "Descrizione_Spese_Imbarco", 
            "BL", "Free_Time", "Validità", "Note", "Origine"
        ])

def salva_database(df):
    df.to_csv(DB_FILE, index=False)

df_master = carica_database()

st.title("🚢 Sistema Nazionale Tariffario Noli Marittimi")
st.write("Console aziendale per l'importazione automatica delle matrici d'offerta Export e la scomposizione delle valute.")

tab_ricerca, tab_automatico, tab_spese_porto, tab_manuale_singolo, tab_database = st.tabs([
    "🔍 Ricerca Tariffe", 
    "📂 1. Carica Matrice Excel Vettori", 
    "✍️ 2. Gestione Spese per Porto e Trade",
    "➕ 3. Inserimento Manuale Spot",
    "📊 Archivio Database Completo"
])

# ==========================================
# TAB 1: INTERFACCIA DI RICERCA
# ==========================================
with tab_ricerca:
    st.header("Consultazione Tariffe Strutturate")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        elenco_pol = sorted(df_master["POL"].dropna().unique()) if not df_master.empty else []
        pol_scelto = st.selectbox("Porto di Partenza (POL)", [""] + [p for p in elenco_pol if p != "SCONOSCIUTO"])
    with col2:
        elenco_pod = sorted(df_master["POD"].dropna().unique()) if not df_master.empty else []
        pod_scelto = st.selectbox("Porto di Destinazione (POD)", [""] + [p for p in elenco_pod if p != "SCONOSCIUTO"])
    with col3:
        elenco_trade = sorted(df_master["Trade"].dropna().unique()) if not df_master.empty else []
        trade_scelto = st.selectbox("Seleziona Trade (Opzionale)", ["TUTTI"] + [t for t in elenco_trade if t != ""])
    with col4:
        tipo_container = st.selectbox("Tipo Container", ["20FT", "40FT", "40HC"])
        
    if pol_scelto and pod_scelto:
        condizione_ricerca = (df_master["POL"] == pol_scelto) & (df_master["POD"] == pod_scelto) & (df_master["Container"] == tipo_container)
        if trade_scelto != "TUTTI":
            condizione_ricerca = condizione_ricerca & (df_master["Trade"] == trade_scelto)
            
        risultati = df_master[condizione_ricerca]
        if not risultati.empty:
            st.success("Tariffe individuate:")
            tabella_visiva = []
            for _, r in risultati.iterrows():
                sym_nolo = "$" if r["Valuta_Nolo"] == "USD" else "€"
                sym_add = "$" if r["Valuta_Addizionali"] == "USD" else "€"
                sym_imb = "$" if r["Valuta_Spese_Imbarco"] == "USD" else "€"
                totale_str = f"{sym_nolo} {r['Totale_Nolo']:.2f}" if r["Valuta_Nolo"] == r["Valuta_Addizionali"] else f"{sym_nolo} {r['Nolo']:.2f} + {sym_add} {r['Addizionali']:.2f}"
                
                tabella_visiva.append({
                    "Compagnia": r["Compagnia"], "Trade": r["Trade"], "POL (Partenza)": r["POL"], "POD (Destinazione)": r["POD"],
                    "Nolo Base": f"{sym_nolo} {r['Nolo']:.2f}", "Totale Addizionali": f"{sym_add} {r['Addizionali']:.2f}",
                    "Dettaglio Add.": r["Descrizione_Addizionali"], "TOTALE NOLO": totale_str,
                    "Spese Imbarco": f"{sym_imb} {r['Spese_Imbarco']:.2f}", "Dettaglio Imbarco": r["Descrizione_Spese_Imbarco"],
                    "Costo BL": f"€ {r['BL']:.2f}", "Free Time": r["Free_Time"], "Validità": r["Validità"], "Note": r["Note"]
                })
            st.dataframe(pd.DataFrame(tabella_visiva), use_container_width=True)
        else:
            st.warning("Nessuna tariffa corrispondente trovata.")

# ==========================================
# TAB 2: MULTI-PARSER RIGIDO (SOLUZIONE BLINDATA)
# ==========================================
with tab_automatico:
    st.header("Estrazione Intelligente con Parser Dedicati")
    col_a1, col_a2 = st.columns(2)
    with col_a1:
        compagnia_file = st.selectbox("Seleziona il Vettore", COMPAGNIE_SUPPORTATE, key="comp_auto")
        trade_file = st.selectbox("Seleziona l'ambito Trade del listino", TRADE_SUPPORTATI, key="trade_auto")
    with col_a2:
        validita_foglio = st.text_input("Validità Temporale Foglio", "01/05/2026-31/05/2026", key="val_auto")
        valuta_matrice = st.radio("Seleziona la valuta dei noli base:", ["USD ($)", "EUR (€)"], horizontal=True)
        valuta_matrice_std = "USD" if "USD" in valuta_matrice else "EUR"
    
    file_caricato = st.file_uploader("Trascina qui il file Excel (.xlsx o .xls)", type=["xlsx", "xls"])
    
    if file_caricato is not None:
        try:
            raw_df = pd.read_excel(file_caricato, header=None)
            st.success("File Excel caricato correttamente in memoria.")
            
            if st.button("Estrai Solo Noli Base"):
                lista_tariffe = []
                
                # ----------------------------------------------------------------------
                # PARSER LAYOUT VERTICALE EXPORT (MSC IPBC, RED SEA, EAF)
                # ----------------------------------------------------------------------
                if trade_file in ["IPBC", "RED SEA", "EAF"]:
                    riga_prezzi_start = None
                    for idx, row in raw_df.iterrows():
                        cella_a_pulita = normalizza_stringa_porto(row.iloc[0])
                        if any(citta in cella_a_pulita for citta in CITTA_ITALIANE_RILEVAMENTO):
                            riga_prezzi_start = idx
                            break
                    
                    if riga_prezzi_start is None:
                        st.error("Impossibile trovare la riga di partenza dei porti italiani nella colonna A.")
                    else:
                        # Estrazione e ffill dei porti esteri POD in alto (riga prezzi - 2)
                        riga_pod_raw = raw_df.iloc[riga_prezzi_start - 2].tolist()
                        riga_pod_pulita = []
                        ultimo_pod_valido = "SCONOSCIUTO"
                        for v in riga_pod_raw:
                            v_str = str(v).strip().upper()
                            if pd.notna(v) and v_str != "" and v_str != "NAN" and "ITALY" not in v_str and "P.O.D." not in v_str and "GUIDELINE" not in v_str:
                                ultimo_pod_valido = " ".join(v_str.split())
                            riga_pod_pulita.append(ultimo_pod_valido)
                            
                        # Estrazione delle colonne container (20' o 40') (riga prezzi - 1)
                        riga_cont_raw = raw_df.iloc[riga_prezzi_start - 1].tolist()
                        riga_cont_pulita = []
                        for v in riga_cont_raw:
                            c_str = str(v).strip()
                            riga_cont_pulita.append("40FT" if "40" in c_str else "20FT")
                        
                        # Iterazione verticale sulle righe di prezzo reali
                        for idx in range(riga_prezzi_start, len(raw_df)):
                            row = raw_df.iloc[idx]
                            pol_raw = normalizza_stringa_porto(row.iloc[0])
                            
                            # Verifica se la riga appartiene a un porto italiano
                            pol_filtrato = None
                            for citta in CITTA_ITALIANE_RILEVAMENTO:
                                if citta in pol_raw:
                                    pol_filtrato = citta
                                    break
                            if not pol_filtrato: continue
                            
                            for col_idx in range(1, len(row)):
                                prezzo_raw = row.iloc[col_idx]
                                try:
                                    price = float(prezzo_raw)
                                    if pd.isna(price) or price <= 0: continue
                                except:
                                    continue
                                
                                pod = riga_pod_pulita[col_idx]
                                container_std = riga_cont_pulita[col_idx]
                                
                                lista_tariffe.append({
                                    "POL": pol_filtrato, "POD": pod, "Compagnia": compagnia_file, "Trade": trade_file, "Container": container_std,
                                    "Nolo": price, "Valuta_Nolo": valuta_matrice_std,
                                    "Addizionali": 0.0, "Valuta_Addizionali": valuta_matrice_std, "Descrizione_Addizionali": "Nessuna surcharge", 
                                    "Totale_Nolo": price, "Spese_Imbarco": 0.0, "Valuta_Spese_Imbarco": "EUR", "Descrizione_Spese_Imbarco": "Nessuna spesa locale", 
                                    "BL": 0.0, "Free_Time": "", "Validità": validita_foglio, "Note": f"Listino Verticale {trade_file}", "Origine": "Automatico"
                                })
                                
                # ----------------------------------------------------------------------
                # PARSER LAYOUT ORIZZONTALE INVERTITO (MIDDLE EAST, FAR EAST, YML, CMA)
                # ----------------------------------------------------------------------
                else:
                    riga_container_idx = None
                    for idx, row in raw_df.iterrows():
                        valori_testo = [str(v).strip().upper() for v in row.values if pd.notna(v)]
                        if any("20DC" in s or "20FT" in s or "20GP" in s or "20'DC" in s for s in valori_testo):
                            riga_container_idx = idx
                            break
                    if riga_container_idx is None: riga_container_idx = 2
                    
                    riga_input_raw = raw_df.iloc[riga_container_idx - 1].tolist()
                    riga_porti_alta = []
                    ultimo_porto_valido = "SCONOSCIUTO"
                    for v in riga_input_raw:
                        val_str = str(v).strip().upper()
                        if pd.notna(v) and val_str != "" and val_str != "NAN" and "ITALY" not in val_str:
                            # Estrae solo il nome pulito se presente nelle nostre chiavi target
                            trovato = False
                            for citta in CITTA_ITALIANE_RILEVAMENTO:
                                if citta in val_str:
                                    ultimo_porto_valido = citta
                                    trovato = True
                                    break
                            if not trovato: ultimo_porto_valido = " ".join(val_str.split())
                        riga_porti_alta.append(ultimo_porto_valido)
                        
                    riga_cont_pulita = [str(v).strip().upper() for v in raw_df.iloc[riga_container_idx]]
                    dati_prezzi = raw_df.iloc[riga_container_idx + 1:].copy()
                    
                    for _, row in dati_prezzi.iterrows():
                        if len(row.values) == 0: continue
                        val_cella_a = str(row.iloc[0]).strip().upper()
                        if val_cella_a in ["", "CURRENCY", "PORT", "TOTAL", "NAN", "SCONOSCIUTO"] or "ALL CARGO" in val_cella_a: continue
                        pod = " ".join(val_cella_a.split())
                        
                        for col_idx in range(1, len(row)):
                            prezzo_raw = row.iloc[col_idx]
                            try:
                                price = float(prezzo_raw)
                                if pd.isna(price) or price <= 0: continue
                            except:
                                continue
                            
                            pol = riga_porti_alta[col_idx]
                            tipo_c_raw = riga_cont_pulita[col_idx]
                            container_std = "20FT" if "20" in tipo_c_raw else ("40HC" if "HQ" in tipo_c_raw or "HC" in tipo_c_raw or "40" in tipo_c_raw else "40FT")
                            
                            lista_tariffe.append({
                                "POL": pol, "POD": pod, "Compagnia": compagnia_file, "Trade": trade_file, "Container": container_std,
                                "Nolo": price, "Valuta_Nolo": valuta_matrice_std,
                                "Addizionali": 0.0, "Valuta_Addizionali": valuta_matrice_std, "Descrizione_Addizionali": "Nessuna surcharge", 
                                "Totale_Nolo": price, "Spese_Imbarco": 0.0, "Valuta_Spese_Imbarco": "EUR", "Descrizione_Spese_Imbarco": "Nessuna spesa locale", 
                                "BL": 0.0, "Free_Time": "", "Validità": validita_foglio, "Note": f"Listino Orizzontale {trade_file}", "Origine": "Automatico"
                            })
                
                df_nuovo = pd.DataFrame(lista_tariffe)
                if not df_nuovo.empty:
                    df_pulito = df_master[(df_master["Compagnia"] != compagnia_file) | (df_master["Trade"] != trade_file)]
                    df_finale = pd.concat([df_pulito, df_nuovo], ignore_index=True)
                    salva_database(df_finale)
                    st.success(f"Estrazione completata con successo! Caricati {len(df_nuovo)} record puliti per {compagnia_file} - Trade {trade_file}.")
                    st.rerun()
                else:
                    st.error("Nessun prezzo valido rilevato. Verifica i parametri selezionati.")
        except Exception as e:
            st.error(f"Errore tecnico durante l'analisi: {e}")

# ==========================================
# TAB 3: GESTIONE SPESE PORTO E TRADE
# ==========================================
with tab_spese_porto:
    st.header("✍️ Inserimento Spese Correttive per Porto e Trade")
    if not df_master.empty:
        col_g1, col_g2 = st.columns(2)
        with col_g1:
            lista_pol_esistenti = [p for p in sorted(df_master["POL"].unique()) if str(p).strip() != "" and p != "SCONOSCIUTO"]
            pol_selezionato_spese = st.selectbox("Seleziona il Porto di Partenza (POL)", lista_pol_esistenti, key="pol_sp_tr")
        with col_g2:
            lista_trade_esistenti = [t for t in sorted(df_master["Trade"].unique()) if str(t).strip() != ""]
            trade_selezionato_spese = st.selectbox("Seleziona il Trade di riferimento", lista_trade_esistenti, key="trade_sp_tr")
            
        st.markdown("---")
        col_an1, col_an2, col_an3 = st.columns(3)
        with col_an1:
            st.subheader("🏢 Spese Imbarco Locali (Valore Fisso)")
            curr_imb = st.radio("Valuta Spese Imbarco:", ["EUR (€)", "USD ($)"], key="c_imb", horizontal=True)
            curr_imb_std = "EUR" if "EUR" in curr_imb else "USD"
            v_thc = st.number_input("THC (base 20FT)", min_value=0.0, step=5.0)
            v_isps = st.number_input("ISPS (base 20FT)", min_value=0.0, step=1.0)
            v_lilo = st.number_input("LILO (base 20FT)", min_value=0.0, step=5.0)
            totale_imb_calcolato = v_thc + v_isps + v_lilo
            st.metric("Totale Imbarco", f"{curr_imb_std} {totale_imb_calcolato:.2f}")
        with col_an2:
            st.subheader("📈 Surcharges / Addizionali (Raddoppia per 40')")
            curr_add = st.radio("Valuta Addizionali:", ["USD ($)", "EUR (€)"], key="c_add", horizontal=True)
            curr_add_std = "USD" if "USD" in curr_add else "EUR"
            v_efs = st.number_input("EFS (base 20FT)", min_value=0.0, step=5.0)
            v_brc = st.number_input("BRC (base 20FT)", min_value=0.0, step=5.0)
            v_eca = st.number_input("ECA (base 20FT)", min_value=0.0, step=5.0)
            v_ets = st.number_input("ETS (base 20FT)", min_value=0.0, step=1.0)
            v_feu = st.number_input("FEU (base 20FT)", min_value=0.0, step=5.0)
            totale_add_calcolato = v_efs + v_brc + v_eca + v_ets + v_feu
            st.metric("Totale Addizionali 20FT", f"{curr_add_std} {totale_add_calcolato:.2f}")
            st.metric("Totale Addizionali 40FT/HC (x2)", f"{curr_add_std} {totale_add_calcolato * 2:.2f}")
        with col_an3:
            st.subheader("📄 Costo Spese Documentali")
            v_bl = st.number_input("Costo Documento BL (€)", min_value=0.0, step=5.0)
            val_free_time = st.text_input("Free Time (Giorni det/dem)", "14 giorni free")
            val_note_libere = st.text_input("Note specifiche rotta", "Valido per nolo in vigore")
            
        if st.button("Calcola Automatismi e Applica"):
            df_modificato = df_master.copy()
            for tipo_c in ["20FT", "40FT", "40HC"]:
                condizione = (df_modificato["POL"] == pol_selezionato_spese) & (df_modificato["Trade"] == trade_selezionato_spese) & (df_modificato["Container"] == tipo_c)
                if not df_modificato[condizione].empty:
                    moltiplicatore_add = 1.0 if tipo_c == "20FT" else 2.0
                    add_riga = totale_add_calcolato * moltiplicatore_add
                    testo_imb = f"THC:{v_thc} | ISPS:{v_isps} | LILO:{v_lilo}"
                    
                    lista_add_descr = []
                    if v_efs > 0: lista_add_descr.append(f"
