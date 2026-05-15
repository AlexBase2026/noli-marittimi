import streamlit as st
import pandas as pd
import os

# Configurazione iniziale dell'interfaccia web aziendale
st.set_page_config(page_title="Tariffario Noli Marittimi", layout="wide", page_icon="🚢")

DB_FILE = "database_noli_ibrido_v2.csv"

def carica_database():
    if os.path.exists(DB_FILE):
        df = pd.read_csv(DB_FILE)
        df['Validità'] = df['Validità'].astype(str).fillna("")
        for col in ["Nolo", "Addizionali", "Totale_Nolo", "Spese_Imbarco"]:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)
        return df
    else:
        return pd.DataFrame(columns=[
            "POL", "POD", "Compagnia", "Container", 
            "Nolo", "Addizionali", "Descrizione_Addizionali", "Totale_Nolo", 
            "Spese_Imbarco", "Descrizione_Spese_Imbarco", "Free_Time", 
            "Validità", "Note", "Origine"
        ])

def salva_database(df):
    df.to_csv(DB_FILE, index=False)

df_master = carica_database()

# --- INTERFACCIA GRAFICA ---
st.title("🚢 Sistema Integrato Gestione Noli Marittimi")
st.write("Applicazione aziendale centralizzata per l'analisi, l'importazione e la modifica delle tariffe di spedizione.")

# Definizione dei 5 Tab (Ripristinato Inserimento Manuale Singolo)
tab_ricerca, tab_automatico, tab_spese_porto, tab_manuale_singolo, tab_database = st.tabs([
    "🔍 Ricerca Tariffe", 
    "📂 1. Carica Matrice Excel", 
    "✍️ 2. Gestione Spese per Porto",
    "➕ 3. Inserimento Manuale Singolo",
    "📊 Archivio Database Completo"
])

# ==========================================
# TAB 1: INTERFACCIA DI RICERCA
# ==========================================
with tab_ricerca:
    st.header("Consultazione Tariffe Condivise")
    col1, col2, col3 = st.columns(3)
    with col1:
        elenco_pol = sorted(df_master["POL"].dropna().unique()) if not df_master.empty else []
        pol_scelto = st.selectbox("Porto di Partenza (POL)", [""] + list(elenco_pol))
    with col2:
        elenco_pod = sorted(df_master["POD"].dropna().unique()) if not df_master.empty else []
        pod_scelto = st.selectbox("Porto di Destinazione (POD)", [""] + list(elenco_pod))
    with col3:
        tipo_container = st.selectbox("Tipo Container", ["20FT", "40FT", "40HC"])
        
    if pol_scelto and pod_scelto:
        risultati = df_master[
            (df_master["POL"] == pol_scelto) & 
            (df_master["POD"] == pod_scelto) & 
            (df_master["Container"] == tipo_container)
        ]
        if not risultati.empty:
            st.success("Tariffe individuate (ordinate per Totale Nolo crescente):")
            mostra_tabella = risultati[[
                "Compagnia", "Nolo", "Addizionali", "Descrizione_Addizionali", 
                "Totale_Nolo", "Spese_Imbarco", "Descrizione_Spese_Imbarco", 
                "Free_Time", "Validità", "Note"
            ]].copy()
            st.dataframe(
                mostra_tabella.style.format({
                    "Nolo": "€ {:.2f}", "Addizionali": "€ {:.2f}",
                    "Totale_Nolo": "€ {:.2f}", "Spese_Imbarco": "€ {:.2f}"
                }), use_container_width=True
            )
        else:
            st.warning("Nessuna tariffa corrispondente trovata.")

# ==========================================
# TAB 2: CARICAMENTO ESCLUSIVO DEI NOLI BASE
# ==========================================
with tab_automatico:
    st.header("Estrazione Automatica Noli Base da Matrice")
    st.write("Carica i listini a griglia. Questo modulo estrarrà i prezzi base azzerando spese e addizionali.")
    
    compagnia_file = st.selectbox("Compagnia Marittima", ["MSC", "CMA", "HAPAG"])
    validita_foglio = st.text_input("Validità Temporale Foglio", "01/05/2026-31/05/2026", key="val_auto")
    file_caricato = st.file_uploader("Trascina il file Excel (.xlsx)", type=["xlsx", "xls"])
    
    if file_caricato is not None:
        try:
            raw_df = pd.read_excel(file_caricato, header=None)
            st.success("File caricato correttamente.")
            
            if st.button("Estrai Solo Noli Base"):
                riga_container_idx = None
                for idx, row in raw_df.iterrows():
                    valori_testo = [str(v).strip() for v in row.values if pd.notna(v)]
                    if any("20" in s for s in valori_testo) and any("40" in s for s in valori_testo):
                        riga_container_idx = idx
                        break
                
                if riga_container_idx is None:
                    riga_container_idx = 2
                
                riga_pod_raw = raw_df.iloc[riga_container_idx - 1].copy()
                riga_pod_pulita = []
                ultimo_pod_valido = "SCONOSCIUTO"
                for v in riga_pod_raw:
                    if pd.notna(v) and str(v).strip() != "" and "ITALY" not in str(v).upper() and "CURRENCY" not in str(v).upper():
                        ultimo_pod_valido = str(v).strip().upper()
                    riga_pod_pulita.append(ultimo_pod_valido)
                
                riga_cont_pulita = [str(v).strip().upper() for v in raw_df.iloc[riga_container_idx]]
                dati_prezzi = raw_df.iloc[riga_container_idx + 1:].copy()
                lista_tariffe = []
                
                for _, row in dati_prezzi.iterrows():
                    pol_raw = row.iloc[0]
                    if pd.isna(pol_raw):
                        continue
                    pol = str(pol_raw).strip().upper()
                    if pol == "" or "CURRENCY" in pol or "MSC" in pol or "PORT" in pol or "MANGALORE" in pol or "ALL ABOVE" in pol:
                        continue
                    
                    for col_idx in range(1, len(row)):
                        prezzo_raw = row.iloc[col_idx]
                        try:
                            prezzo = float(prezzo_raw)
                            if pd.isna(prezzo) or prezzo <= 0:
                                continue
                        except:
                            continue
                        
                        pod = riga_pod_pulita[col_idx]
                        tipo_c_raw = riga_cont_pulita[col_idx]
                        tipi_da_generare = ["20FT"] if "20" in tipo_c_raw else ["40FT", "40HC"]
                        
                        if "(" in pod:
                            pod = pod.split("(")[0].strip()
                            
                        for container_std in tipi_da_generare:
                            lista_tariffe.append({
                                "POL": pol, "POD": pod, "Compagnia": compagnia_file, "Container": container_std,
                                "Nolo": prezzo, "Addizionali": 0.0, "Descrizione_Addizionali": "", 
                                "Totale_Nolo": prezzo, "Spese_Imbarco": 0.0, "Descrizione_Spese_Imbarco": "", 
                                "Free_Time": "", "Validità": validita_foglio, "Note": "Importato da matrice", "Origine": "Automatico"
                            })
                
                df_nuovo_standard = pd.DataFrame(lista_tariffe)
                if not df_nuovo_standard.empty:
                    df_pulito_precedente = df_master[df_master["Compagnia"] != compagnia_file]
                    df_finale = pd.concat([df_pulito_precedente, df_nuovo_standard], ignore_index=True)
                    salva_database(df_finale)
                    st.success(f"Estrazione completata! Caricati {len(df_nuovo_standard)} noli base puri.")
                    st.rerun()
        except Exception as e:
            st.error(f"Errore durante l'estrazione: {e}")

# ==========================================
# TAB 3: AGGIORNAMENTO DI MASSA DELLE SPESE PER POL
# ==========================================
with tab_spese_porto:
    st.header("✍️ Gestione Spese e Addizionali Localizzate per Porto")
    st.write("Aggiorna massivamente le spese locali e le addizionali per tutte le rotte che partono dallo stesso POL.")
    
    if not df_master.empty:
        lista_pol_esistenti = sorted(df_master["POL"].unique())
        pol_selezionato_spese = st.selectbox("Scegli il porto di partenza (POL) da valorizzare", lista_pol_esistenti)
        container_selezionato_spese = st.selectbox("Seleziona tipo container per l'aggiornamento di massa", ["20FT", "40FT", "40HC", "TUTTI"])
        
        col_s1, col_s2 = st.columns(2)
        with col_s1:
            val_spese_imb = st.number_input("Spese Imbarco (€/$)", min_value=0.0, step=5.0, key="sp_imb_mas")
            desc_spese_imb = st.text_input("Descrizione Spese Imbarco", "THC + ISPS locale", key="desc_imb_mas")
            val_free_time = st.text_input("Free Time dedicato", "14 giorni det/dem", key="ft_mas")
        with col_s2:
            val_addizionali = st.number_input("Addizionali Correttive (€/$)", min_value=0.0, step=5.0, key="add_mas")
            desc_addizionali = st.text_input("Descrizione Addizionali", "EFS + BRC", key="desc_add_mas")
            val_note_libere = st.text_input("Note specifiche per questo porto", "Soggetto a variazioni", key="note_mas")
            
        if st.button(f"Applica Costi Massivi a {pol_selezionato_spese}"):
            df_modificato = df_master.copy()
            condizione = (df_modificato["POL"] == pol_selezionato_spese)
            if container_selezionato_spese != "TUTTI":
                condizione = condizione & (df_modificato["Container"] == container_selezionato_spese)
                
            if not df_modificato[condizione].empty:
                df_modificato.loc[condizione, "Spese_Imbarco"] = val_spese_imb
                df_modificato.loc[condizione, "Descrizione_Spese_Imbarco"] = desc_spese_imb
                df_modificato.loc[condizione, "Addizionali"] = val_addizionali
                df_modificato.loc[condizione, "Descrizione_Addizionali"] = desc_addizionali
                df_modificato.loc[condizione, "Free_Time"] = val_free_time
                df_modificato.loc[condizione, "Note"] = val_note_libere
                df_modificato.loc[condizione, "Totale_Nolo"] = df_modificato.loc[condizione, "Nolo"] + val_addizionali
                
                salva_database(df_modificato)
                st.success(f"Aggiornamento di massa completato per {pol_selezionato_spese}!")
                st.rerun()
    else:
        st.info("Archivio vuoto. Carica prima un listino o fai un inserimento singolo.")

# ==========================================
# TAB 4: RIPRISTINATO - INSERIMENTO MANUALE SINGOLO / QUOTAZIONI SPOT
# ==========================================
with tab_manuale_singolo:
    st.header("➕ Inserimento Manuale Singola Tariffa Spot")
    st.write("Usa questo modulo per registrare una tariffa singola non presente nei listini automatici o per aggiungere un porto nuovo.")
    
    with st.form("Form Inserimento Singolo"):
        col_m1, col_m2, col_m3 = st.columns(3)
        with col_m1:
            man_pol = st.text_input("Porto POL (Partenza)").upper().strip()
            man_carrier = st.text_input("Compagnia (es. MSC)").upper().strip()
            man_nolo = st.number_input("Nolo Base (€)", min_value=0.0, step=50.0, key="nolo_sin")
            man_spese = st.number_input("Spese Imbarco (€)", min_value=0.0, step=10.0, key="spese_sin")
            man_desc_spese = st.text_input("Descrizione Spese Imbarco", key="desc_spese_sin")
        with col_m2:
            man_pod = st.text_input("Porto POD (Destinazione)").upper().strip()
            man_container = st.selectbox("Tipo Container", ["20FT", "40FT", "40HC"], key="cont_sin")
            man_addizionali = st.number_input("Addizionali (€)", min_value=0.0, step=10.0, key="add_sin")
            man_desc_addizionali = st.text_input("Descrizione Add.li", key="desc_add_sin")
            man_freetime = st.text_input("Free Time", key="ft_sin")
        with col_m3:
            man_validita = st.text_input("Data Validità", "01/05/2026-31/05/2026", key="val_sin")
            man_note = st.text_area("Campo Note Libero", height=150, key="note_sin")
            
        st.write(" ")
        st.write(" ")
        invia_form = st.form_submit_button("Salva Singola Tariffa Spot")
        
        if invia_form:
            if man_pol and man_pod and man_carrier:
                totale_calcolato = man_nolo + man_addizionali
                nuova_riga = pd.DataFrame([{
                    "POL": man_pol, "POD": man_pod, "Compagnia": man_carrier, "Container": man_container,
                    "Nolo": man_nolo, "Addizionali": man_addizionali, "Descrizione_Addizionali": man_desc_addizionali,
                    "Totale_Nolo": totale_calcolato, "Spese_Imbarco": man_spese, "Descrizione_Spese_Imbarco": man_desc_spese,
                    "Free_Time": man_freetime, "Validità": man_validita, "Note": man_note, "Origine": "Manuale"
                }])
                df_aggiornato = pd.concat([df_master, nuova_riga], ignore_index=True)
                salva_database(df_aggiornato)
                st.success(f"Tariffa salvata! Aggiunta correttamente la tratta {man_pol} ➡️ {man_pod}.")
                st.rerun()
            else:
                st.error("I campi POL, POD e Compagnia sono obbligatori per il salvataggio.")

# ==========================================
# TAB 5: ARCHIVIO E MANUTENZIONE DATABASE
# ==========================================
with tab_database:
    st.header("Visualizzazione Tabellare di Controllo")
    st.dataframe(df_master, use_container_width=True)
    if st.button("🗑️ Svuota Intero Database"):
        if os.path.exists(DB_FILE):
            os.remove(DB_FILE)
        st.rerun()
