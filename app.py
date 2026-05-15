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

# Dizionario per tradurre le stringhe di intestazione MSC nei porti commerciali puliti
DIZIONARIO_PORTI_FISSI = {
    "ITGOA": "GENOVA", "ITLIV": "LIVORNO", "ITSPE": "LA SPEZIA", "ITVCE": "VENEZIA", "ITNAP": "NAPOLI", "ITAO1": "ANCONA",
    "GENOVA, LA SPEZIA, TRIESTE, AND GIOIA TAURO": "GENOVA",
    "LA SPEZIA, GENOVA, GIOIA TAURO, TRIESTE": "GENOVA",
    "VENEZIA, RAVENNA, ANCONA": "VENEZIA",
    "NAPOLI/SALERNO (VIA NOLO)": "NAPOLI",
    "BARI, CIVITAVECCHIA, TRAPANI, POZZALLO, AUGUSTA & TERMINI IMERESE": "BARI",
    "BARI, CIVITAVECCHIA, PALERMO, TRAPANI, POZZALLO, AUGUSTA & TERMINI IMERESE": "BARI",
    "CAGLIARI": "CAGLIARI"
}

def normalizza_porto_msc(valore_cella):
    testo = str(valore_cella).strip().replace("\n", " ")
    if not testo or testo.upper() in ["NAN", "NONE", "", "0", "0.0"]:
        return "SCONOSCIUTO"
    testo = " ".join(testo.split())
    testo_upper = testo.upper()
    if testo_upper in DIZIONARIO_PORTI_FISSI:
        return DIZIONARIO_PORTI_FISSI[testo_upper]
    return testo

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
                    "Nolo Base": f"{sym_nolo} {r['Nolo']:.2f}", "Totale Addizionali": f"{sym_add} {r['Addizionali']:.2f}", "Dettaglio Addizionali": r["Descrizione_Addizionali"],
                    "TOTALE NOLO BASE+ADD": totale_str, "Totale Spese Imbarco": f"{sym_imb} {r['Spese_Imbarco']:.2f}", "Dettaglio Imbarco": r["Descrizione_Spese_Imbarco"],
                    "Costo BL": f"€ {r['BL']:.2f}", "Free Time": r["Free_Time"], "Validità": r["Validità"], "Note": r["Note"]
                })
            st.dataframe(pd.DataFrame(tabella_visiva), use_container_width=True)
        else:
            st.warning("Nessuna tariffa corrispondente trovata.")

# ==========================================
# TAB 2: CARICAMENTO MATRICE EXCEL MULTI-PARSER
# ==========================================
with tab_automatico:
    st.header("Estrazione Intelligente con Parser Dedicati")
    col_a1, col_a2 = st.columns(2)
    with col_a1:
        compagnia_file = st.selectbox("Seleziona il Vettore", COMPAGNIE_SUPPORTATE, key="comp_auto")
        trade_file = st.selectbox("Seleziona l'ambito Trade del listino", TRADE_SUPPORTATI, key="trade_auto")
    with col_a2:
        validita_foglio = st.text_input("Validità Temporale Foglio", "01/05/2026-31/05/2026", key="val_auto")
        valuta_matrice = st.radio("Seleziona la valuta dei noli base della griglia Excel:", ["USD ($)", "EUR (€)"], horizontal=True)
        valuta_matrice_std = "USD" if "USD" in valuta_matrice else "EUR"
    
    file_caricato = st.file_uploader("Trascina qui il file Excel (.xlsx o .xls)", type=["xlsx", "xls"])
    
    if file_caricato is not None:
        try:
            raw_df = pd.read_excel(file_caricato, header=None)
            st.success("File Excel caricato correttamente.")
            
            if st.button("Estrai Solo Noli Base"):
                lista_tariffe = []
                
                # --- PARSER LAYOUT 2: MATRICI ORIZZONTALI (CMA, COSCO, EVERGREEN, YML, ONE, MAERSK) ---
                if compagnia_file in ["CMA", "COSCO", "EVERGREEN", "MAERSK", "ONE", "YML"]:
                    riga_container_idx = None
                    for idx, row in raw_df.iterrows():
                        valori_testo = [str(v).strip().upper() for v in row.values if pd.notna(v)]
                        if any("20DC" in s or "20FT" in s or "20GP" in s or "20'DC" in s or "20DC" == s for s in valori_testo):
                            riga_container_idx = idx
                            break
                    
                    if riga_container_idx is None: riga_container_idx = 2
                    
                    riga_input_raw = raw_df.iloc[riga_container_idx - 1].tolist()
                    riga_porti_alta = []
                    ultimo_porto_valido = "SCONOSCIUTO"
                    for v in riga_input_raw:
                        val_str = str(v).strip()
                        if pd.notna(v) and val_str != "" and val_str.upper() != "NAN" and "ITALY" not in val_str.upper():
                            ultimo_porto_valido = normalizza_porto_msc(v)
                        riga_porti_alta.append(ultimo_porto_valido)
                    
                    riga_cont_pulita = [str(v).strip().upper() for v in raw_df.iloc[riga_container_idx]]
                    dati_prezzi = raw_df.iloc[riga_container_idx + 1:].copy()
                    
                    for _, row in dati_prezzi.iterrows():
                        if len(row.values) == 0: continue
                        pol_cella = row.iloc[0]
                        if pd.isna(pol_cella) or str(pol_cella).strip() == "": continue
                        
                        pod = str(pol_cella).strip().upper()
                        if pod in ["", "CURRENCY", "PORT", "TOTAL", "NAN", "SCONOSCIUTO"]: continue
                        
                        for col_idx in range(1, len(row)):
                            prezzo_raw = row.iloc[col_idx]
                            try:
                                price = float(prezzo_raw)
                                if pd.isna(price) or price <= 0: continue
                            except:
                                continue
                            
                            pol = riga_porti_alta[col_idx]
                            tipo_c_raw = riga_cont_pulita[col_idx]
                            container_std = "20FT" if "20" in tipo_c_raw else ("40HC" if "HQ" in tipo_c_raw or "HC" in tipo_c_raw else "40FT")
                            
                            lista_tariffe.append({
                                "POL": pol, "POD": pod, "Compagnia": compagnia_file, "Trade": trade_file, "Container": container_std,
                                "Nolo": price, "Valuta_Nolo": valuta_matrice_std,
                                "Addizionali": 0.0, "Valuta_Addizionali": valuta_matrice_std, "Descrizione_Addizionali": "Nessuna surcharge", 
                                "Totale_Nolo": price, "Spese_Imbarco": 0.0, "Valuta_Spese_Imbarco": "EUR", "Descrizione_Spese_Imbarco": "Nessuna spesa locale", 
                                "BL": 0.0, "Free_Time": "", "Validità": validita_foglio, "Note": "Layout Orizzontale Export", "Origine": "Automatico"
                            })
                            
                # --- PARSER LAYOUT 1: STRUTTURA VERTICALE STANDARD INTANGIBILE (MSC, HAPAG, ECC.) ---
                else:
                    riga_container_idx = None
                    for idx, row in raw_df.iterrows():
                        valori_testo = [str(v).strip() for v in row.values if pd.notna(v)]
                        if any("20" in s for s in valori_testo) and any("40" in s for s in valori_testo):
                            riga_container_idx = idx
                            break
                    
                    if riga_container_idx is None: riga_container_idx = 2
                    
                    riga_pod_raw = raw_df.iloc[riga_container_idx - 1].tolist()
                    riga_pod_pulita = []
                    ultimo_pod_valido = "SCONOSCIUTO"
                    
                    for v in riga_pod_raw:
                        val_str = str(v).strip()
                        if pd.notna(v) and val_str != "" and val_str.upper() != "NAN":
                            ultimo_pod_valido = normalizza_porto_msc(v)
                        riga_pod_pulita.append(ultimo_pod_valido)
                    
                    riga_cont_pulita = [str(v).strip().upper() for v in raw_df.iloc[riga_container_idx]]
                    dati_prezzi = raw_df.iloc[riga_container_idx + 1:].copy()
                    
                    for _, row in dati_prezzi.iterrows():
                        if len(row.values) == 0: continue
                        pol_cella = row.iloc[0]
                        if pd.isna(pol_cella) or str(pol_cella).strip() == "": continue
                        
                        testo_fianco = str(pol_cella).strip().upper()
                        if testo_fianco in ["", "CURRENCY", "PORT", "TOTAL", "GUIDELINE"]: continue
                        
                        for col_idx in range(1, len(row)):
                            prezzo_raw = row.iloc[col_idx]
                            try:
                                price = float(prezzo_raw)
                                if pd.isna(price) or price <= 0: continue
                            except:
                                continue
                            
                            elemento_alto = riga_pod_pulita[col_idx]
                            tipo_c_raw = riga_cont_pulita[col_idx]
                            container_std = "20FT" if "20" in tipo_c_raw else "40FT"
                            
                            # LOGICA DI INVERSIONE AUTOMATICA EXPORT MIDDLE EAST:
                            # Se l'elemento in alto è un porto italiano convertito (quindi non è SCONOSCIUTO), invertiamo i ruoli
                            if elemento_alto != "SCONOSCIUTO":
                                pol_finale = elemento_alto
                                pod_finale = testo_fianco
                            else:
                                pol_finale = testo_fianco
                                pod_finale = elemento_alto
                                
                            lista_tariffe.append({
                                "POL": pol_finale, "POD": pod_finale, "Compagnia": compagnia_file, "Trade": trade_file, "Container": container_std,
                                "Nolo": price, "Valuta_Nolo": valuta_matrice_std,
                                "Addizionali": 0.0, "Valuta_Addizionali": valuta_matrice_std, "Descrizione_Addizionali": "Nessuna surcharge", 
                                "Totale_Nolo": price, "Spese_Imbarco": 0.0, "Valuta_Spese_Imbarco": "EUR", "Descrizione_Spese_Imbarco": "Nessuna spesa locale", 
                                "BL": 0.0, "Free_Time": "", "Validità": validita_foglio, "Note": f"Layout MSC {trade_file}", "Origine": "Automatico"
                            })
                
                df_nuovo = pd.DataFrame(lista_tariffe)
                if not df_nuovo.empty:
                    df_pulito = df_master[(df_master["Compagnia"] != compagnia_file) | (df_master["Trade"] != trade_file)]
                    df_finale = pd.concat([df_pulito, df_nuovo], ignore_index=True)
                    salva_database(df_finale)
                    st.success(f"Estrazione completata! Caricati {len(df_nuovo)} noli.")
                    st.rerun()
                else:
                    st.error("Nessun prezzo rilevato. Verifica il formato.")
        except Exception as e:
            st.error(f"Errore tecnico durante l'analisi: {e}")

# ==========================================
# TAB 3: GESTIONE SPESE PORTO E TRADE CONGIUNTI
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
            st.write(" ")
            st.metric("Totale Imbarco (20FT / 40FT / 40HC)", f"{curr_imb_std} {totale_imb_calcolato:.2f}")
        with col_an2:
            st.subheader("📈 Surcharges / Addizionali (Raddoppia per 40')")
            curr_add = st.radio("Valuta Addizionali / Surcharges:", ["USD ($)", "EUR (€)"], key="c_add", horizontal=True)
            curr_add_std = "USD" if "USD" in curr_add else "EUR"
            v_efs = st.number_input("EFS (base 20FT)", min_value=0.0, step=5.0)
            v_brc = st.number_input("BRC (base 20FT)", min_value=0.0, step=5.0)
            v_eca = st.number_input("ECA (base 20FT)", min_value=0.0, step=5.0)
            v_ets = st.number_input("ETS (base 20FT)", min_value=0.0, step=1.0)
            v_feu = st.number_input("FEU (base 20FT)", min_value=0.0, step=5.0)
            totale_add_calcolato = v_efs + v_brc + v_eca + v_ets + v_feu
            st.write(" ")
            st.metric("Totale Addizionali 20FT", f"{curr_add_std} {totale_add_calcolato:.2f}")
            st.metric("Totale Addizionali 40FT / 40HC (Automatico x2)", f"{curr_add_std} {totale_add_calcolato * 2:.2f}")
        with col_an3:
            st.subheader("📄 Costo Spese Documentali")
            v_bl = st.number_input("Costo Documento BL (€)", min_value=0.0, step=5.0)
            st.write(" ")
            val_free_time = st.text_input("Free Time (Giorni det/dem)", "14 giorni free")
            val_note_libere = st.text_input("Note specifiche rotta", "Valido per nolo in vigore")
            
        if st.button(f"Calcola Automatismi e Applica a {pol_selezionato_spese} - Trade {trade_selezionato_spese}"):
            df_modificato = df_master.copy()
            for tipo_c in ["20FT", "40FT", "40HC"]:
                condizione = (df_modificato["POL"] == pol_selezionato_spese) & (df_modificato["Trade"] == trade_selezionato_spese) & (df_modificato["Container"] == tipo_c)
                if not df_modificato[condizione].empty:
                    moltiplicatore_add = 1.0 if tipo_c == "20FT" else 2.0
                    imb_riga = totale_imb_calcolato
                    add_riga = totale_add_calcolato * moltiplicatore_add
                    testo_imb = f"THC:{v_thc} | ISPS:{v_isps} | LILO:{v_lilo}"
                    
                    lista_add_descr = []
                    if v_efs > 0: lista_add_descr.append(f"EFS:{v_efs * moltiplicatore_add}")
                    if v_brc > 0: lista_add_descr.append(f"BRC:{v_brc * moltiplicatore_add}")
                    if v_eca > 0: lista_add_descr.append(f"ECA:{v_eca * moltiplicatore_add}")
                    if v_ets > 0: lista_add_descr.append(f"ETS:{v_ets * moltiplicatore_add}")
                    if v_feu > 0: lista_add_descr.append(f"FEU:{v_feu * moltiplicatore_add}")
                    testo_add = " + ".join(lista_add_descr) if lista_add_descr else "Nessuna surcharge"
                    
                    df_modificato.loc[condizione, "Spese_Imbarco"] = float(imb_riga)
                    df_modificato.loc[condizione, "Descrizione_Spese_Imbarco"] = str(testo_imb)
                    df_modificato.loc[condizione, "Valuta_Spese_Imbarco"] = str(curr_imb_std)
                    df_modificato.loc[condizione, "Addizionali"] = float(add_riga)
                    df_modificato.loc[condizione, "Descrizione_Addizionali"] = str(testo_add)
                    df_modificato.loc[condizione, "Valuta_Addizionali"] = str(curr_add_std)
                    df_modificato.loc[condizione, "BL"] = float(v_bl)
                    df_modificato.loc[condizione, "Free_Time"] = str(val_free_time)
                    df_modificato.loc[condizione, "Note"] = str(val_note_libere)
                    df_modificato.loc[condizione, "Totale_Nolo"] = df_modificato.loc[condizione, "Nolo"] + float(add_riga)
            salva_database(df_modificato)
            st.success(f"Configurazione completata per {pol_selezionato_spese}!")
            st.rerun()
    else:
        st.info("Nessun dato di nolo base presente. Esegui prima l'importazione nel Tab 1.")

# ==========================================
# TAB 4: INSERIMENTO MANUALE SPOT
# ==========================================
with tab_manuale_singolo:
    st.header("➕ Inserimento Manuale Singola Tariffa Spot Scomposta")
    with st.form("Form Inserimento Analitico"):
        col_m1, col_m2, col_m3 = st.columns(3)
        with col_m1:
            man_pol = st.text_input("Porto POL (Partenza)").upper().strip()
            man_carrier = st.selectbox("Seleziona Compagnia", COMPAGNIE_SUPPORTATE, key="comp_man")
            man_trade = st.selectbox("Seleziona Trade", TRADE_SUPPORTATI, key="trade_man")
            man_nolo = st.number_input("Nolo Base", min_value=0.0, step=50.0)
            man_v_nolo = st.selectbox("Valuta Nolo Base", ["USD", "EUR"])
            st.markdown("**Spese Locali Imbarco**")
            man_thc = st.number_input("THC", min_value=0.0, step=5.0)
            man_isps = st.number_input("ISPS", min_value=0.0, step=1.0)
            man_lilo = st.number_input("LILO", min_value=0.0, step=5.0)
            man_v_imb = st.selectbox("Valuta Spese Imbarco", ["EUR", "USD"])
        with col_m2:
            man_pod = st.text_input("Porto POD (Destinazione)").upper().strip()
            man_container = st.selectbox("Tipo Container", ["20FT", "40FT", "40HC"])
            st.markdown("**Surcharges / Addizionali**")
            man_efs = st.number_input("EFS", min_value=0.0, step=5.0)
            man_brc = st.number_input("BRC", min_value=0.0, step=5.0)
            man_eca = st.number_input("ECA", min_value=0.0, step=5.0)
            man_ets = st.number_input("ETS", min_value=0.0, step=1.0)
            man_feu = st.number_input("FEU", min_value=0.0, step=5.0)
            man_v_add = st.selectbox("Valuta Addizionali", ["USD", "EUR"])
        with col_m3:
            st.markdown("**Spese Documentali e Note**")
            man_bl = st.number_input("Costo Polizza BL (€)", min_value=0.0, step=5.0)
            man_freetime = st.text_input("Free Time dedicato")
            man_validita = st.text_input("Data Validità", "01/05/2026-31/05/2026")
            man_note = st.text_area("Campo Note Libero", height=100)
            
        invia_form = st.form_submit_button("Salva Tariffa Spot")
        if invia_form:
            if man_pol and man_pod and man_carrier:
                tot_add_sin = man_efs + man_brc + man_eca + man_ets + man_feu
                tot_imb_sin = man_thc + man_isps + man_lilo
                tot_nolo_sin = man_nolo + tot_add_sin
                testo_add_sin = f"EFS:{man_efs} BRC:{man_brc} ECA:{man_eca} ETS:{man_ets} FEU:{man_feu}"
                testo_imb_sin = f"THC:{man_thc} | ISPS:{man_isps} | LILO:{man_lilo}"
                
                pol_std = normalizza_porto_msc(man_pol)
                pod_std = normalizza_porto_msc(man_pod)
                
                nuova_riga = pd.DataFrame([{
                    "POL": pol_std, "POD": pod_std, "Compagnia": man_carrier, "Trade": man_trade, "Container": man_container,
                    "Nolo": man_nolo, "Valuta_Nolo": man_v_nolo, "Addizionali": tot_add_sin, "Valuta_Addizionali": man_v_add, "Descrizione_Addizionali": str(testo_add_sin),
                    "Totale_Nolo": tot_nolo_sin, "Spese_Imbarco": tot_imb_sin, "Valuta_Spese_Imbarco": man_v_imb, "Descrizione_Spese_Imbarco": str(testo_imb_sin),
                    "BL": man_bl, "Free_Time": str(man_freetime), "Validità": str(man_validita), "Note": str(man_note), "Origine": "Manuale"
                }])
                salva_database(pd.concat([df_master, nuova_riga], ignore_index=True))
                st.success("Tariffa spot salvata correttamente!")
                st.rerun()

with tab_database:
    st.header("Visualizzazione Tabellare di Controllo (Dati nel CSV)")
    st.dataframe(df_master, use_container_width=True)
    if st.button("🗑 Svuota Intero Database"):
        if os.path.exists(DB_FILE):
            os.remove(DB_FILE)
        st.rerun()
