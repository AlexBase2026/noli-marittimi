import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Tariffario Noli Marittimi", layout="wide", page_icon="🚢")

DB_FILE = "database_noli_analitico_valute.csv"

COMPAGNIE_SUPPORTATE = [
    "MSC", "CMA", "HAPAG", "MAERSK", "EVERGREEN", "MESSINA", "GRIMALDI", 
    "COSCO", "YML", "ONE", "AKKON", "ARKAS", "TARROS", "HMM", "OOCL", 
    "SAHEL", "SCI LINE", "ZIM", "COSIARMA", "MARFRET", "BORCHARD", 
    "COTUNAV", "MAGUISA"
]

# Elenco dei Trade/Direttrici per mappare regole e layout differenti all'interno della stessa compagnia
TRADE_SUPPORTATI = [
    "IPBC (India / Pakistan Subcontinent)",
    "Middle & Far East / Red Sea (Export Griglia Orizzontale)",
    "Standard Trade / Altri Layout"
]

DIZIONARIO_PORTI_FISSI = {
    "ITGOA": "GENOVA", "ITLIV": "LIVORNO", "ITSPE": "LA SPEZIA", 
    "ITVCE": "VENEZIA", "ITNAP": "NAPOLI", "ITAO1": "ANCONA"
}

def normalizza_porto_msc(valore_cella):
    """Mantiene il testo del porto esattamente come scritto nel listino, senza troncarlo"""
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
        for col_testo in ["POL", "POD", "Compagnia", "Container", "Descrizione_Addizionali", "Descrizione_Spese_Imbarco", "Free_Time", "Validità", "Note", "Origine", "Valuta_Nolo", "Valuta_Addizionali", "Valuta_Spese_Imbarco", "Trade"]:
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

st.title("🚢 Sistema Multi-Parser e Multi-Trade Noli Marittimi")

tab_ricerca, tab_automatico, tab_spese_porto, tab_manuale_singolo, tab_database = st.tabs([
    "🔍 Ricerca Tariffe", 
    "📂 1. Carica Matrice Excel Vettori", 
    "✍️ 2. Gestione Spese per Porto",
    "➕ 3. Inserimento Manuale Spot",
    "📊 Archivio Database Completo"
])

# ==========================================
# TAB 1: INTERFACCIA DI RICERCA
# ==========================================
with tab_ricerca:
    st.header("Consultazione Tariffe e Scomposizione Valute")
    col1, col2, col3 = st.columns(3)
    with col1:
        elenco_pol = sorted(df_master["POL"].dropna().unique()) if not df_master.empty else []
        pol_scelto = st.selectbox("Porto di Partenza (POL)", [""] + [p for p in elenco_pol if p != "SCONOSCIUTO"])
    with col2:
        elenco_pod = sorted(df_master["POD"].dropna().unique()) if not df_master.empty else []
        pod_scelto = st.selectbox("Porto di Destinazione (POD)", [""] + [p for p in elenco_pod if p != "SCONOSCIUTO"])
    with col3:
        tipo_container = st.selectbox("Tipo Container", ["20FT", "40FT", "40HC"])
        
    if pol_scelto and pod_scelto:
        risultati = df_master[
            (df_master["POL"] == pol_scelto) & 
            (df_master["POD"] == pod_scelto) & 
            (df_master["Container"] == tipo_container)
        ]
        if not risultati.empty:
            st.success("Tariffe individuate:")
            
            tabella_visiva = []
            for _, r in risultati.iterrows():
                sym_nolo = "$" if r["Valuta_Nolo"] == "USD" else "€"
                sym_add = "$" if r["Valuta_Addizionali"] == "USD" else "€"
                sym_imb = "$" if r["Valuta_Spese_Imbarco"] == "USD" else "€"
                
                trade_txt = f" [{r['Trade']}]" if "Trade" in r and r["Trade"] != "" else ""
                
                if r["Valuta_Nolo"] == r["Valuta_Addizionali"]:
                    totale_str = f"{sym_nolo} {r['Totale_Nolo']:.2f}"
                else:
                    totale_str = f"{sym_nolo} {r['Nolo']:.2f} + {sym_add} {r['Addizionali']:.2f}"
                
                tabella_visiva.append({
                    "Compagnia": f"{r['Compagnia']}{trade_txt}",
                    "POL (Partenza)": r["POL"],
                    "POD (Destinazione)": r["POD"],
                    "Nolo Base": f"{sym_nolo} {r['Nolo']:.2f}",
                    "Totale Addizionali": f"{sym_add} {r['Addizionali']:.2f}",
                    "Dettaglio Addizionali": r["Descrizione_Addizionali"],
                    "TOTALE NOLO BASE+ADD": totale_str,
                    "Totale Spese Imbarco": f"{sym_imb} {r['Spese_Imbarco']:.2f}",
                    "Dettaglio Imbarco (Local)": r["Descrizione_Spese_Imbarco"],
                    "Costo BL": f"€ {r['BL']:.2f}",
                    "Free Time": r["Free_Time"],
                    "Validità": r["Validità"],
                    "Note": r["Note"]
                })
            
            st.dataframe(pd.DataFrame(tabella_visiva), use_container_width=True)
        else:
            st.warning("Nessuna tariffa corrispondente trovata.")

# ==========================================
# TAB 2: ARCHITETTURA MULTI-TRADE (STRUTTURA AGGIORNATA)
# ==========================================
with tab_automatico:
    st.header("Estrazione Intelligente con Parser Dedicati")
    
    col_l1, col_l2 = st.columns(2)
    with col_l1:
        compagnia_file = st.selectbox("Seleziona il Vettore del listino", COMPAGNIE_SUPPORTATE, key="comp_auto")
        validita_foglio = st.text_input("Validità Temporale Foglio", "01/05/2026-31/05/2026", key="val_auto")
    with col_l2:
        # NUOVO CAMPO: Consente di indicare la direttrice esatta per scegliere la scomposizione geometrica corretta
        trade_file = st.selectbox("Seleziona la Direttrice / Trade del listino", TRADE_SUPPORTATI, key="trade_auto")
        valuta_matrice = st.radio("Seleziona la valuta dei noli base della griglia Excel:", ["USD ($)", "EUR (€)"], horizontal=True)
        valuta_matrice_std = "USD" if "USD" in valuta_matrice else "EUR"
    
    file_caricato = st.file_uploader("Trascina il file Excel della compagnia", type=["xlsx", "xls"])
    
    if file_caricato is not None:
        try:
            raw_df = pd.read_excel(file_caricato, header=None)
            st.success("File Excel caricato correttamente in memoria.")
            
            if st.button("Estrai Solo Noli Base"):
                lista_tariffe = []
                
                # --------------------------------------------------------
                # PARSER LAYOUT 3: MSC FAR EAST / RED SEA (NUOVO LAYOUT EXPORT RILEVATO)
                # --------------------------------------------------------
                if compagnia_file == "MSC" and "Middle & Far East" in trade_file:
                    riga_container_idx = None
                    for idx, row in raw_df.iterrows():
                        valori_testo = [str(v).strip().upper() for v in row.values if pd.notna(v)]
                        # Cerca la riga di sbarramento dei container 20 e 40
                        if any(s == "20'" or s == "20" or s == "20FT" for s in valori_testo):
                            riga_container_idx = idx
                            break
                    
                    if riga_container_idx is None: riga_container_idx = 13 # Fallback basato sulla foto
                    
                    # Legge la riga superiore dei macro-porti italiani (POL)
                    riga_pol_raw = raw_df.iloc[riga_container_idx - 1].tolist()
                    riga_pol_mappati = []
                    ultimo_pol_valido = "SCONOSCIUTO"
                    
                    for v in riga_pol_raw:
                        val_str = str(v).strip()
                        if pd.notna(v) and val_str != "" and val_str.upper() != "NAN" and "P.O.D." not in val_str.upper() and "GUIDELINE" not in val_str.upper():
                            ultimo_pol_valido = val_str
                        riga_pol_mappati.append(ultimo_pol_valido)
                        
                    riga_cont_pulita = [str(v).strip().upper() for v in raw_df.iloc[riga_container_idx]]
                    dati_prezzi = raw_df.iloc[riga_container_idx + 1:].copy()
                    
                    for _, row in dati_prezzi.iterrows():
                        if len(row.values) == 0: continue
                        # In questo layout il porto estero di destinazione (POD) è nella colonna 0 in verticale
                        pod_cella = row.iloc[0]
                        if pd.isna(pod_cella) or str(pod_cella).strip() == "" or "P.O.D." in str(pod_cella).upper() or "FAR EAST" in str(pod_cella).upper(): 
                            continue
                        pod = normalizza_porto_msc(pod_cella)
                        
                        for col_idx in range(1, len(row)):
                            prezzo_raw = row.iloc[col_idx]
                            try:
                                prezzo = float(prezzo_raw)
                                if pd.isna(prezzo) or prezzo <= 0: continue
                            except:
                                continue
                            
                            pol_blocco_testo = riga_pol_mappati[col_idx]
                            tipo_c_raw = riga_cont_pulita[col_idx]
                            
                            # Logica di scompattamento delle celle dei porti italiani (es. "Venezia, Ravenna, Ancona")
                            porti_italiani_scompattati = []
                            for p_singolo in pol_blocco_testo.replace("and", ",").split(","):
                                nome_p_pulito = p_singolo.strip().upper()
                                if nome_p_pulito != "":
                                    porti_italiani_scompattati.append(nome_p_pulito)
                                    
                            if not porti_italiani_scompattati:
                                porti_italiani_scompattati = ["GENOVA"]
                                
                            tipi_da_generare = ["20FT"] if "20" in tipo_c_raw else ["40FT", "40HC"]
                            
                            # Genera combinazioni pulite per ogni porto del blocco e ogni container richiesto
                            for pol_final in porti_italiani_scompattati:
                                for container_std in tipi_da_generare:
                                    lista_tariffe.append({
                                        "POL": pol_final, "POD": pod, "Compagnia": compagnia_file, "Trade": "FAR EAST", "Container": container_std,
                                        "Nolo": prezzo, "Valuta_Nolo": valuta_matrice_std,
                                        "Addizionali": 0.0, "Valuta_Addizionali": valuta_matrice_std, "Descrizione_Addizionali": "Nessuna surcharge", 
                                        "Totale_Nolo": prezzo, "Spese_Imbarco": 0.0, "Valuta_Spese_Imbarco": "EUR", "Descrizione_Spese_Imbarco": "Nessuna spesa locale", 
                                        "BL": 0.0, "Free_Time": "", "Validità": validita_foglio, "Note": "Importato da Matrice Far East Export", "Origine": "Automatico"
                                    })

                # --------------------------------------------------------
                # PARSER LAYOUT 2: MATRICI ORIZZONTALI (CMA, COSCO, EVERGREEN, YML, ONE, MAERSK)
                # --------------------------------------------------------
                elif compagnia_file in ["CMA", "COSCO", "EVERGREEN", "MAERSK", "ONE", "YML"] or "Subcontinent" not in trade_file:
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
                        
                        pod = normalizza_porto_msc(pol_cella)
                        if pod in ["", "CURRENCY", "PORT", "TOTAL", "NAN", "SCONOSCIUTO"]: continue
                        
                        for col_idx in range(1, len(row)):
                            prezzo_raw = row.iloc[col_idx]
                            try:
                                prezzo = float(prezzo_raw)
                                if pd.isna(prezzo) or prezzo <= 0: continue
                            except:
                                continue
                            
                            pol = riga_porti_alta[col_idx]
                            tipo_c_raw = riga_cont_pulita[col_idx]
                            
                            if "20" in tipo_c_raw: container_std = "20FT"
                            elif "HQ" in tipo_c_raw or "HC" in tipo_c_raw or "40HQ" in tipo_c_raw or "40HC" in tipo_c_raw: container_std = "40HC"
                            else: container_std = "40FT"
                                
                            lista_tariffe.append({
                                "POL": pol, "POD": pod, "Compagnia": compagnia_file, "Trade": "EXPORT TRADE", "Container": container_std,
                                "Nolo": prezzo, "Valuta_Nolo": valuta_matrice_std,
                                "Addizionali": 0.0, "Valuta_Addizionali": valuta_matrice_std, "Descrizione_Addizionali": "Nessuna surcharge", 
                                "Totale_Nolo": prezzo, "Spese_Imbarco": 0.0, "Valuta_Spese_Imbarco": "EUR", "Descrizione_Spese_Imbarco": "Nessuna spesa locale", 
                                "BL": 0.0, "Free_Time": "", "Validità": validita_foglio, "Note": f"Importato da layout Orizzontale", "Origine": "Automatico"
                            })
                            
                # --------------------------------------------------------
                # PARSER LAYOUT 1: STRUTTURA IPBC STANDARD (IL PRIMO FILE MSC)
                # --------------------------------------------------------
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
                        if pd.notna(v) and val_str != "" and val_str.upper() != "NAN" and "ITALY" not in val_str.upper():
                            ultimo_pod_valido = normalizza_porto_msc(v)
                        riga_pod_pulita.append(ultimo_pod_valido)
                    
                    riga_cont_pulita = [str(v).strip().upper() for v in raw_df.iloc[riga_container_idx]]
                    dati_prezzi = raw_df.iloc[riga_container_idx + 1:].copy()
                    
                    for _, row in dati_prezzi.iterrows():
                        if len(row.values) == 0: continue
                        pol_cella = row.iloc[0]
                        if pd.isna(pol_cella) or str(pol_cella).strip() == "": continue
                        
                        pol = normalizza_porto_msc(pol_cella)
                        if pol in ["", "CURRENCY", "PORT", "TOTAL", "SCONOSCIUTO"]: continue
                        
                        for col_idx in range(1, len(row)):
                            prezzo_raw = row.iloc[col_idx]
                            try:
                                prezzo = float(prezzo_raw)
                                if pd.isna(prezzo) or prezzo <= 0: continue
                            except:
                                continue
                            
                            pod = riga_pod_pulita[col_idx]
                            tipo_c_raw = riga_cont_pulita[col_idx]
                            tipi_da_generare = ["20FT"] if "20" in tipo_c_raw else ["40FT", "40HC"]
                            
                            for container_std in tipi_da_generare:
                                lista_tariffe.append({
                                    "POL": pol, "POD": pod, "Compagnia": compagnia_file, "Trade": "IPBC", "Container": container_std,
                                    "Nolo": prezzo, "Valuta_Nolo": valuta_matrice_std,
                                    "Addizionali": 0.0, "Valuta_Addizionali": valuta_matrice_std, "Descrizione_Addizionali": "Nessuna surcharge", 
                                    "Totale_Nolo": prezzo, "Spese_Imbarco": 0.0, "Valuta_Spese_Imbarco": "EUR", "Descrizione_Spese_Imbarco": "Nessuna spesa locale", 
                                    "BL": 0.0, "Free_Time": "", "Validità": validita_foglio, "Note": f"Importato da layout IPBC", "Origine": "Automatico"
                                })
                
                df_nuovo_standard = pd.DataFrame(lista_tariffe)
                if not df_nuovo_standard.empty:
                    # Rimuove solo i record vecchi della stessa compagnia e dello stesso specifico trade per non fare sovrascritture incrociate
                    trade_tag = "FAR EAST" if "Far East" in trade_file else ("IPBC" if "IPBC" in trade_file else "EXPORT TRADE")
                    df_pulito_precedente = df_master[~((df_master["Compagnia"] == compagnia_file) & (df_master["Trade"] == trade_tag))]
                    df_finale = pd.concat([df_pulito_precedente, df_nuovo_standard], ignore_index=True)
                    salva_database(df_finale)
                    st.success(f"Estrazione completata! Mappate correttamente {len(df_nuovo_standard)} rotte commerciali per MSC {trade_tag}.")
                    st.rerun()
                else:
                    st.error("Nessun prezzo rilevato. Verifica la formattazione dei campi.")
        except Exception as e:
            st.error(f"Errore tecnico: {e}")

# ==========================================
# TAB 3: GESTIONE SPESE PORTO MOLTIPLICATORI
# ==========================================
with tab_spese_porto:
    st.header("✍️ Inserimento Spese per Porto")
    if not df_master.empty:
        lista_pol_esistenti = [p for p in sorted(df_master["POL"].unique()) if str(p).strip() != "" and p != "SCONOSCIUTO"]
        pol_selezionato_spese = st.selectbox("Seleziona il Porto di Partenza (POL) da valorizzare", lista_pol_esistenti)
        
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
            
        if st.button(f"Calcola Automatismi e Applica a {pol_selezionato_spese}"):
            df_modificato = df_master.copy()
            for tipo_c in ["20FT", "40FT", "40HC"]:
                condizione = (df_modificato["POL"] == pol_selezionato_spese) & (df_modificato["Container"] == tipo_c)
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
            st.success("Database aggiornato applicando i totali!")
            st.rerun()

with tab_manuale_singolo:
    st.header("➕ Inserimento Manuale Singola Tariffa Spot Scomposta")
    with st.form("Form Inserimento Analitico"):
        col_m1, col_m2, col_m3 = st.columns(3)
        with col_m1:
            man_pol = st.text_input("Porto POL (Partenza)").upper().strip()
            man_carrier = st.selectbox("Seleziona Compagnia", COMPAGNIE_SUPPORTATE, key="comp_man")
            man_trade = st.text_input("Trade / Direttrice (es. RED SEA)", "FAR EAST")
            man_nolo = st.number_input("Nolo
