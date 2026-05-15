import streamlit as st
import pandas as pd
import os

# Configurazione iniziale dell'interfaccia web aziendale
st.set_page_config(page_title="Tariffario Noli Marittimi", layout="wide", page_icon="🚢")

DB_FILE = "database_noli_matrice.csv"

def carica_database():
    if os.path.exists(DB_FILE):
        df = pd.read_csv(DB_FILE)
        df['Validità'] = df['Validità'].astype(str).fillna("")
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

# --- INTERFACCIA ---
st.title("🚢 Sistema Condiviso Ricerca Noli Marittimi")

tab_ricerca, tab_automatico, tab_manuale, tab_database = st.tabs([
    "🔍 Ricerca Tariffe", "📂 Carica File Compagnie", "✍️ Inserimento Manuale", "📊 Visualizza Intero Database"
])

# ==========================================
# TAB 1: INTERFACCIA DI RICERCA
# ==========================================
with tab_ricerca:
    st.header("Cerca la migliore quotazione")
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
            st.success("Opzioni tariffarie trovate:")
            st.dataframe(risultati.sort_values(by="Totale_Nolo"), use_container_width=True)
        else:
            st.warning("Nessuna tariffa trovata per questa selezione.")

# ==========================================
# TAB 2: PARSER MATRICI EXCEL AGGIORNATO E ROBUSTO
# ==========================================
with tab_automatico:
    st.header("Caricamento Listini Matrix Excel")
    st.write("Modulo ottimizzato per griglie MSC con porti sia nelle righe che nelle colonne.")
    
    compagnia_file = st.selectbox("Seleziona la compagnia", ["MSC", "CMA", "HAPAG"])
    validita_foglio = st.text_input("Periodo di Validità per questo listino", "01/05/2026-31/05/2026")
    
    file_caricato = st.file_uploader("Trascina qui il file Excel", type=["xlsx", "xls"])
    
    if file_caricato is not None:
        try:
            # Legge l'excel senza intestazioni per gestire le celle unite liberamente
            raw_df = pd.read_excel(file_caricato, header=None)
            st.success("File Excel caricato in memoria.")
            
            if st.button("Avvia Conversione Automatica della Griglia"):
                # 1. Trova la riga dei container (20', 40') in modo sicuro convertendo tutto in testo
                riga_container_idx = None
                for idx, row in raw_df.iterrows():
                    valori_testo = [str(v).strip() for v in row.values if pd.notna(v)]
                    # Se la riga contiene elementi riconducibili a 20 e 40, è la riga di sbarramento
                    if any("20" in s for s in valori_testo) and any("40" in s for s in valori_testo):
                        riga_container_idx = idx
                        break
                
                if riga_container_idx is None:
                    riga_container_idx = 2  # Fallback standard
                
                # 2. Ricostruisce la riga dei POD sopra a quella dei container riempiendo i vuoti delle celle unite
                riga_pod_raw = raw_df.iloc[riga_container_idx - 1].copy()
                riga_pod_pulita = []
                ultimo_pod_valido = "SCONOSCIUTO"
                
                for v in riga_pod_raw:
                    if pd.notna(v) and str(v).strip() != "" and "ITALY" not in str(v).upper():
                        ultimo_pod_valido = str(v).strip().upper()
                    riga_pod_pulita.append(ultimo_pod_valido)
                
                # 3. Legge la riga dei tipi container
                riga_cont_pulita = [str(v).strip().upper() for v in raw_df.iloc[riga_container_idx]]
                
                # 4. Estrae i prezzi puri a partire dalla riga successiva
                dati_prezzi = raw_df.iloc[riga_container_idx + 1:].copy()
                lista_tariffe_standardizzate = []
                
                for _, row in dati_prezzi.iterrows():
                    # Il primo elemento della riga è sempre il porto di partenza (POL)
                    pol_raw = row.iloc[0]
                    if pd.isna(pol_raw):
                        continue
                        
                    pol = str(pol_raw).strip().upper()
                    # Salta righe di intestazione secondarie o note in fondo al file
                    if pol == "" or "CURRENCY" in pol or "MSC" in pol or "PORT" in pol:
                        continue
                    
                    # Scorri i prezzi presenti nelle colonne successive
                    for col_idx in range(1, len(row)):
                        prezzo_raw = row.iloc[col_idx]
                        
                        # Verifica se la cella contiene un prezzo numerico valido
                        try:
                            prezzo = float(prezzo_raw)
                            if pd.isna(prezzo) or prezzo <= 0:
                                continue
                        except:
                            continue
                        
                        # Recupera POD e Tipo Container corrispondenti alla colonna corrente
                        pod = riga_pod_pulita[col_idx]
                        tipo_c_raw = riga_cont_pulita[col_idx]
                        
                        if "20" in tipo_c_raw:
                            container_std = "20FT"
                        elif "40HC" in tipo_c_raw or "HC" in tipo_c_raw:
                            container_std = "40HC"
                        else:
                            container_std = "40FT"
                        
                        # Pulisce eventuali note tra parentesi nel nome del porto di destinazione
                        if "(" in pod:
                            pod = pod.split("(")[0].strip()
                            
                        lista_tariffe_standardizzate.append({
                            "POL": pol, "POD": pod, "Compagnia": compagnia_file, "Container": container_std,
                            "Nolo": prezzo, "Addizionali": 0.0, "Descrizione_Addizionali": "Incluso nel file matrice",
                            "Totale_Nolo": prezzo, "Spese_Imbarco": 0.0, "Descrizione_Spese_Imbarco": "",
                            "Free_Time": "Vedi listino", "Validità": validita_foglio, "Note": "Importato da matrice", "Origine": "Automatico"
                        })
                
                df_nuovo_standard = pd.DataFrame(lista_tariffe_standardizzate)
                
                if not df_nuovo_standard.empty:
                    df_pulito_precedente = df_master[df_master["Compagnia"] != compagnia_file]
                    df_finale = pd.concat([df_pulito_precedente, df_nuovo_standard], ignore_index=True)
                    salva_database(df_finale)
                    
                    st.success(f"Conversione completata con successo! Generate {len(df_nuovo_standard)} combinazioni POL/POD.")
                    st.rerun()
                else:
                    st.error("Nessun dato tariffario estratto. Verifica la struttura delle celle.")
        except Exception as e:
            st.error(f"Errore tecnico durante la conversione: {e}")

# ==========================================
# TAB 3: INSERIMENTO MANUALE
# ==========================================
with tab_manuale:
    st.header("Inserisci una singola tariffa")
    with st.form("Form Inserimento Dettagliato"):
        col_m1, col_m2, col_m3 = st.columns(3)
        with col_m1:
            man_pol = st.text_input("Porto POL (Partenza)").upper().strip()
            man_carrier = st.text_input("Compagnia (es. MSC)").upper().strip()
            man_nolo = st.number_input("Nolo Base (€)", min_value=0.0, step=50.0)
            man_spese = st.number_input("Spese Imbarco (€)", min_value=0.0, step=10.0)
            man_desc_spese = st.text_input("Descrizione Spese Imbarco")
        with col_m2:
            man_pod = st.text_input("Porto POD (Destinazione)").upper().strip()
            man_container = st.selectbox("Tipo Container", ["20FT", "40FT", "40HC"])
            man_addizionali = st.number_input("Addizionali (€)", min_value=0.0, step=10.0)
            man_desc_addizionali = st.text_input("Descrizione Add.li")
            man_freetime = st.text_input("Free Time")
        with col_m3:
            man_validita = st.text_input("Data Validità (es. 01/05/2026-31/05/2026)", "01/05/2026-31/05/2026")
            man_note = st.text_area("Campo Note Libero", height=150)
            
        invia_form = st.form_submit_button("Salva Tariffa nel Database")
        if invia_form:
            if man_pol and man_pod and man_carrier:
                totale_calcolato = man_nolo + man_addizionali
                nuova_riga = pd.DataFrame([{
                    "POL": man_pol, "POD": man_pod, "Compagnia": man_carrier, "Container": man_container,
                    "Nolo": man_nolo, "Addizionali": man_addizionali, "Descrizione_Addizionali": man_desc_addizionali,
                    "Totale_Nolo": totale_calcolato, "Spese_Imbarco": man_spese, "Descrizione_Spese_Imbarco": man_desc_spese,
                    "Free_Time": man_freetime, "Validità": man_validita, "Note": man_note, "Origine": "Manuale"
                }])
                df_master = pd.concat([df_master, nuova_riga], ignore_index=True)
                salva_database(df_master)
                st.success("Salvataggio completato!")
                st.rerun()

with tab_database:
    st.header("Archivio Storico")
    st.dataframe(df_master, use_container_width=True)
