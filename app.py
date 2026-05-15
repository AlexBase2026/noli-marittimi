import streamlit as st
import pandas as pd
import os
from datetime import datetime

# Configurazione iniziale della pagina web aziendale
st.set_page_config(page_title="Tariffario Noli Marittimi", layout="wide", page_icon="🚢")

DB_FILE = "database_noli.csv"

# Funzione per caricare il database centrale
def carica_database():
    if os.path.exists(DB_FILE):
        df = pd.read_csv(DB_FILE)
        df['Validità'] = pd.to_datetime(df['Validità'], errors='coerce').dt.date
        return df
    else:
        return pd.DataFrame(columns=["POL", "POD", "Compagnia", "20FT", "40FT", "40HC", "Validità", "Origine"])

# Funzione per salvare i dati nel database
def salva_database(df):
    df.to_csv(DB_FILE, index=False)

df_master = carica_database()

# --- INTERFACCIA GRAFICA ---
st.title("🚢 Sistema Condiviso Ricerca Noli Marittimi")
st.write("Strumento aziendale centralizzato per la consultazione e gestione delle tariffe spot.")

# Creazione dei Tab per separare le funzioni
tab_ricerca, tab_automatico, tab_manuale, tab_database = st.tabs([
    "🔍 Ricerca Tariffe", 
    "📂 Carica File Compagnie", 
    "✍️ Inserimento Manuale",
    "📊 Visualizza Intero Database"
])

# ==========================================
# TAB 1: INTERFACCIA DI RICERCA PER I COLLEGHI
# ==========================================
with tab_ricerca:
    st.header("Cerca la migliore quotazione")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        elenco_pol = sorted(df_master["POL"].dropna().unique()) if not df_master.empty else []
        pol_scelto = st.selectbox("Porto di Partenza (POL)", [""] + list(elenco_pol), index=0)
        
    with col2:
        elenco_pod = sorted(df_master["POD"].dropna().unique()) if not df_master.empty else []
        pod_scelto = st.selectbox("Porto di Destinazione (POD)", [""] + list(elenco_pod), index=0)
        
    with col3:
        tipo_container = st.selectbox("Tipo Container", ["20FT", "40FT", "40HC"])

    if pol_scelto and p_scelto :
        # Filtra i dati in base alle scelte dell'utente
        oggi = datetime.now().date()
        risultati = df_master[
            (df_master["POL"] == pol_scelto) & 
            (df_master["POD"] == pod_scelto) & 
            (df_master["Validità"] >= oggi)  # Mostra solo tariffe non scadute
        ]
        
        if not risultati.empty:
            st.success(f"Trovate {len(risultati)} opzioni valide da {pol_scelto} a {pod_scelto}:")
            
            # Ordina dal prezzo più basso a quello più alto per il container scelto
            risultati_ordinati = risultati.sort_values(by=tipo_container)
            
            # Formatta la visualizzazione della tabella dei prezzi
            mostra_tabella = risultati_ordinati[["Compagnia", tipo_container, "Validità", "Origine"]].copy()
            st.dataframe(
                mostra_tabella.style.format({tipo_container: "€ {:.2f}"}),
                use_container_width=True
            )
        else:
            st.warning("Nessuna tariffa attiva trovata per questa tratta. Controlla la scadenza nel database.")
    else:
        st.info("Seleziona un porto di partenza e uno di destinazione per visualizzare i prezzi.")

# ==========================================
# TAB 2: AGGIORNAMENTO AUTOMATICO (CARICAMENTO FILE)
# ==========================================
with tab_automatico:
    st.header("Caricamento listini Excel Compagnie (CMA, MSC, HAPAG, ecc.)")
    st.write("Trascina qui il file Excel ricevuto dalla compagnia. Il sistema lo elaborerà automaticamente.")
    
    compagnia_file = st.selectbox("Seleziona la compagnia del file", ["MSC", "CMA", "HAPAG"])
    file_caricato = st.file_uploader("Scegli un file Excel o CSV", type=["xlsx", "xls", "csv"])
    
    if file_caricato is not None:
        if st.button("Elabora e Unisci Dati"):
            try:
                # Lettura del file (Excel o CSV)
                if file_caricato.name.endswith('.csv'):
                    df_Nuovo = pd.read_csv(file_caricato)
                else:
                    df_Nuovo = pd.read_excel(file_caricato)
                
                # --- LOGICA DI STANDARDIZZAZIONE DEI FILE DELLE COMPAGNIE ---
                # Questo blocco adatta le colonne dei file esterni alla nostra struttura standard.
                # Nota: i nomi delle colonne ("Port_of_Loading", ecc.) vanno adattati a come la compagnia scrive i suoi file.
                
                df_standard = pd.DataFrame()
                
                if compagnia_file == "MSC":
                    df_standard["POL"] = df_Nuovo["Origin_Port"].astype(str).str.upper().str.strip()
                    df_standard["POD"] = df_Nuovo["Destination_Port"].astype(str).str.upper().str.strip()
                    df_standard["20FT"] = pd.to_numeric(df_Nuovo["Rate_20DV"], errors='coerce')
                    df_standard["40FT"] = pd.to_numeric(df_Nuovo["Rate_40DV"], errors='coerce')
                    df_standard["40HC"] = pd.to_numeric(df_Nuovo["Rate_40HC"], errors='coerce')
                    df_standard["Validità"] = pd.to_datetime(df_Nuovo["Expiry_Date"]).dt.date
                    
                elif compagnia_file == "CMA":
                    df_standard["POL"] = df_Nuovo["POL_Code"].astype(str).str.upper().str.strip()
                    df_standard["POD"] = df_Nuovo["POD_Code"].astype(str).str.upper().str.strip()
                    df_standard["20FT"] = pd.to_numeric(df_Nuovo["20GP"], errors='coerce')
                    df_standard["40FT"] = pd.to_numeric(df_Nuovo["40GP"], errors='coerce')
                    df_standard["40HC"] = pd.to_numeric(df_Nuovo["40HC"], errors='coerce')
                    df_standard["Validità"] = pd.to_datetime(df_Nuovo["Valid_To"]).dt.date
                    
                elif compagnia_file == "HAPAG":
                    df_standard["POL"] = df_Nuovo["From"].astype(str).str.upper().str.strip()
                    df_standard["POD"] = df_Nuovo["To"].astype(str).str.upper().str.strip()
                    df_standard["20FT"] = pd.to_numeric(df_Nuovo["20'"], errors='coerce')
                    df_standard["40FT"] = pd.to_numeric(df_Nuovo["40'"], errors='coerce')
                    df_standard["40HC"] = pd.to_numeric(df_Nuovo["40'HC"], errors='coerce')
                    df_standard["Validità"] = pd.to_datetime(df_Nuovo["Expiration"]).dt.date

                df_standard["Compagnia"] = compagnia_file
                df_standard["Origine"] = "Automatico"
                
                # Rimuove righe con dati essenziali mancanti
                df_standard = df_standard.dropna(subset=["POL", "POD"])
                
                # Unione con il vecchio database (rimuovendo i vecchi dati della stessa compagnia per evitare duplicati)
                df_pulito_precedente = df_master[df_master["Compagnia"] != compagnia_file]
                df_finale = pd.concat([df_pulito_precedente, df_standard], ignore_index=True)
                
                salva_database(df_finale)
                st.success(f"Database aggiornato! Importate correttamente {len(df_standard)} tariffe da {compagnia_file}.")
                st.rerun()
                
            except Exception as e:
                st.error(f"Errore durante l'elaborazione. Verifica che i nomi delle colonne del file corrispondano. Dettaglio: {e}")

# ==========================================
# TAB 3: INSERIMENTO MANUALE SPOT QUICK
# ==========================================
with tab_manuale:
    st.header("Inserisci una singola tariffa manualmente")
    with st.form("Form Inserimento"):
        col_m1, col_m2, col_m3 = st.columns(3)
        with col_m1:
            man_pol = st.text_input("Porto POL (es. SHANGHAI)").upper().strip()
            man_20 = st.number_input("Prezzo 20FT (€)", min_value=0.0, step=50.0)
        with col_m2:
            man_pod = st.text_input("Porto POD (es. GENOA)").upper().strip()
            man_40 = st.number_input("Prezzo 40FT (€)", min_value=0.0, step=50.0)
        with col_m3:
            man_carrier = st.text_input("Compagnia (es. MSC)").upper().strip()
            man_40hc = st.number_input("Prezzo 40HC (€)", min_value=0.0, step=50.0)
            
        man_scadenza = st.date_input("Data Scadenza Tariffa")
        
        invia_form = st.form_submit_button("Salva Tariffa nel Database")
        
        if invia_form:
            if man_pol and man_pod and man_carrier:
                nuova_riga = pd.DataFrame([{
                    "POL": man_pol, "POD": man_pod, "Compagnia": man_carrier,
                    "20FT": man_20, "40FT": man_40, "40HC": man_40hc,
                    "Validità": man_scadenza, "Origine": "Manuale"
                }])
                df_aggiornato = pd.concat([df_master, nuova_riga], ignore_index=True)
                salva_database(df_aggiornato)
                st.success("Tariffa inserita con successo!")
                st.rerun()
            else:
                st.error("I campi POL, POD e Compagnia sono obbligatori per il salvataggio.")

# ==========================================
# TAB 4: VISUALIZZAZIONE COMPLETA E MANUTENZIONE
# ==========================================
with tab_database:
    st.header("Archivio Storico Completo")
    if not df_master.empty:
        st.dataframe(df_master, use_container_width=True)
        if st.button("🗑️ Cancella l'intero Database (Reset Totale)"):
            if os.path.exists(DB_FILE):
                os.remove(DB_FILE)
            st.success("Database azzerato.")
            st.rerun()
    else:
        st.info("Il database è attualmente vuoto.")
