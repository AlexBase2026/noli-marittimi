import streamlit as st
import pandas as pd
import os
import numpy as np

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
            st.success(f"Opzioni tariffarie trovate:")
            st.dataframe(risultati.sort_values(by="Totale_Nolo"), use_container_width=True)
        else:
            st.warning("Nessuna tariffa trovata per questa selezione.")

# ==========================================
# TAB 2: PARSER INTELLIGENTE PER MATRICI EXCEL (MSC)
# ==========================================
with tab_automatico:
    st.header("Caricamento Listini Matrix Excel / PDF")
    st.write("Questo modulo elabora automaticamente i file strutturati a matrice (es. griglie MSC con porti in riga e colonna).")
    
    compagnia_file = st.selectbox("Seleziona la compagnia", ["MSC", "CMA", "HAPAG"])
    
    # Campo di input per specificare il periodo di validità di questo intero foglio
    validita_foglio = st.text_input("Periodo di Validità per questo listino (es. 01/05/2026-31/05/2026)", "01/05/2026-31/05/2026")
    
    file_caricato = st.file_uploader("Trascina qui il file Excel o PDF", type=["xlsx", "xls", "pdf"])
    
    if file_caricato is not None:
        if file_caricato.name.endswith('.pdf'):
            st.info("Funzione di lettura testo PDF attiva. Per strutture a griglia complessa come MSC, si raccomanda l'uso del file Excel originale.")
        else:
            try:
                # Carica l'excel leggendo le prime righe per analizzare la struttura
                raw_df = pd.read_excel(file_caricato, header=None)
                
                st.success("File Excel caricato in memoria. Elaborazione della matrice...")
                
                if st.button("Avvia Conversione Automatica della Griglia"):
                    # 1. Trova la riga dove ci sono i tipi di container (20', 40') per capire dove iniziano i dati
                    # Nel file MSC solitamente è la riga che contiene '20'' e '40'' alternati
                    riga_container_idx = None
                    for idx, row in raw_df.iterrows():
                        row_str = row.astype(str).values
                        if any("20'" in s or "20" in s for s in row_str) and any("40'" in s or "40" in s for s in row_str):
                            riga_container_idx = idx
                            break
                    
                    if riga_container_idx is None:
                        # Fallback se non trova le intestazioni esatte
                        riga_container_idx = 2 
                    
                    # Rileva i POD (stanno sopra la riga dei container)
                    riga_pod = raw_df.iloc[riga_container_idx - 1].forward_fill() if riga_container_idx > 0 else raw_df.iloc[0]
                    # Riempi i blocchi uniti dei POD (es. Nhava Sheva copre sia la colonna 20' che 40')
                    riga_pod = pd.Series(riga_pod).ffill().tolist()
                    
                    riga_cont = raw_df.iloc[riga_container_idx].tolist()
                    
                    # Estrae la matrice dei prezzi puri e dei POL
                    dati_prezzi = raw_df.iloc[riga_container_idx + 1:].copy()
                    
                    lista_tariffe_standardizzate = []
                    
                    # Scorri ogni riga di porto di partenza (POL)
                    for _, row in dati_prezzi.iterrows():
                        pol = str(row[0]).strip().upper()
                        if pd.isna(row[0]) or pol == "NAN" or "CURRENCY" in pol or pol == "":
                            continue
                            
                        # Scorri le colonne partendo dalla seconda (indice 1)
                        for col_idx in range(1, len(row)):
                            prezzo_val = row[col_idx]
                            
                            # Trasforma in numero pulito
                            try:
                                prezzo = float(prezzo_val)
                            except:
                                continue
                                
                            if pd.isna(prezzo) or prezzo <= 0:
                                continue
                                
                            # Determina il POD e il tipo di Container associato a questa colonna
                            pod = str(riga_pod[col_idx]).strip().upper()
                            tipo_c_raw = str(riga_cont[col_idx]).strip()
                            
                            # Normalizza il nome del container
                            if "20" in tipo_c_raw:
                                container_std = "20FT"
                            elif "40HC" in tipo_c_raw or "hc" in tipo_c_raw:
                                container_std = "40HC"
                            else:
                                container_std = "40FT"
                                
                            # Rimuove note dai nomi dei porti per pulizia
                            if "(" in pod:
                                pod = pod.split("(")[0].strip()
                                
                            lista_tariffe_standardizzate.append({
                                "POL": pol,
                                "POD": pod,
                                "Compagnia": compagnia_file,
                                "Container": container_std,
                                "Nolo": prezzo,
                                "Addizionali": 0.0,
                                "Descrizione_Addizionali": "Incluso nel file matrice",
                                "Totale_Nolo": prezzo, # Al momento imposta uguale, modificabile da inserimento
                                "Spese_Imbarco": 0.0,
                                "Descrizione_Spese_Imbarco": "",
                                "Free_Time": "Vedi note generali",
                                "Validità": validita_foglio,
                                "Note": f"Importato da griglia Trade {compagnia_file}",
                                "Origine": "Automatico"
                            })
                    
                    df_nuovo_standard = pd.DataFrame(lista_tariffe_standardizzate)
                    
                    if not df_nuovo_standard.empty:
                        # Unione e salvataggio
                        df_pulito_precedente = df_master[df_master["Compagnia"] != compagnia_file]
                        df_finale = pd.concat([df_pulito_precedente, df_nuovo_standard], ignore_index=True)
                        salva_database(df_finale)
                        
                        st.success(f"Conversione completata! Generate correttamente {len(df_nuovo_standard)} combinazioni tariffarie POL/POD per {compagnia_file}.")
                        st.rerun()
                    else:
                        st.error("Nessun prezzo numerico valido estratto. Verifica la struttura del foglio Excel.")
                        
            except Exception as e:
                st.error(f"Errore tecnico durante la conversione della matrice: {e}")

# (I TAB 3 e 4 mantengono le stesse funzioni descrittive e note inserite in precedenza)
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
                salva_database(pd.concat([df_master, nuova_riga], ignore_index=True))
                st.success("Salvataggio completato!")
                st.rerun()

with tab_database:
    st.header("Archivio Storico")
    st.dataframe(df_master, use_container_width=True)
