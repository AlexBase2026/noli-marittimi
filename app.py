import streamlit as st
import pandas as pd
import os

# Configurazione iniziale dell'interfaccia web aziendale
st.set_page_config(page_title="Tariffario Noli Marittimi", layout="wide", page_icon="🚢")

DB_FILE = "database_noli_periodi.csv"

# Funzione per caricare il database centrale con il campo validità testuale
def carica_database():
    if os.path.exists(DB_FILE):
        df = pd.read_csv(DB_FILE)
        # Forza la colonna Validità a testo per gestire i range (es. 01/05/2026-31/05/2026)
        df['Validità'] = df['Validità'].astype(str).fillna("")
        return df
    else:
        return pd.DataFrame(columns=[
            "POL", "POD", "Compagnia", "Container", 
            "Nolo", "Addizionali", "Descrizione_Addizionali", "Totale_Nolo", 
            "Spese_Imbarco", "Descrizione_Spese_Imbarco", "Free_Time", 
            "Validità", "Note", "Origine"
        ])

# Funzione per salvare i dati nel database
def salva_database(df):
    df.to_csv(DB_FILE, index=False)

df_master = carica_database()

# --- INTERFACCIA GRAFICA ---
st.title("🚢 Sistema Condiviso Ricerca Noli Marittimi")
st.write("Strumento aziendale centralizzato per la consultazione delle tariffe con periodi di validità estesi.")

# Creazione dei Tab per separare le funzioni
tab_ricerca, tab_automatico, tab_manuale, tab_database = st.tabs([
    "🔍 Ricerca Tariffe", 
    "📂 Carica File Compagnie", 
    "✍️ Inserimento Manuale",
    "📊 Visualizza Intero Database"
])

# ==========================================
# TAB 1: INTERFACCIA DI RICERCA COMPLETA
# ==========================================
with tab_ricerca:
    st.header("Cerca la migliore quotazione e verifica il periodo di validità")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        elenco_pol = sorted(df_master["POL"].dropna().unique()) if not df_master.empty else []
        pol_scelto = st.selectbox("Porto di Partenza (POL)", [""] + list(elenco_pol), index=0)
        
    with col2:
        elenco_pod = sorted(df_master["POD"].dropna().unique()) if not df_master.empty else []
        pod_scelto = st.selectbox("Porto di Destinazione (POD)", [""] + list(elenco_pod), index=0)
        
    with col3:
        tipo_container = st.selectbox("Tipo Container", ["20FT", "40FT", "40HC"])

    if pol_scelto and pod_scelto:
        # Filtra i dati per rotta e container (senza filtro data automatico per mostrare i range scritti a mano)
        risultati = df_master[
            (df_master["POL"] == pol_scelto) & 
            (df_master["POD"] == pod_scelto) & 
            (df_master["Container"] == tipo_container)
        ]
        
        if not risultati.empty:
            st.success(f"Trovate {len(risultati)} opzioni registrate per container {tipo_container} da {pol_scelto} a {pod_scelto}:")
            
            # Ordina dal Totale Nolo più basso
            risultati_ordinati = risultati.sort_values(by="Totale_Nolo")
            
            # Tabella con colonna Validità visualizzata come testo libero
            mostra_tabella = risultati_ordinati[[
                "Compagnia", "Nolo", "Addizionali", "Descrizione_Addizionali", 
                "Totale_Nolo", "Spese_Imbarco", "Descrizione_Spese_Imbarco", 
                "Free_Time", "Validità", "Note", "Origine"
            ]].copy()
            
            st.dataframe(
                mostra_tabella.style.format({
                    "Nolo": "€ {:.2f}",
                    "Addizionali": "€ {:.2f}",
                    "Totale_Nolo": "€ {:.2f}",
                    "Spese_Imbarco": "€ {:.2f}"
                }),
                use_container_width=True
            )
        else:
            st.warning("Nessuna tariffa trovata per questa combinazione.")
    else:
        st.info("Seleziona un porto di partenza e uno di destinazione per iniziare la ricerca.")

# ==========================================
# TAB 2: AGGIORNAMENTO AUTOMATICO VIA EXCEL
# ==========================================
with tab_automatico:
    st.header("Caricamento listini Excel Compagnie")
    st.write("Il sistema importerà i dati convertendo i campi data nel formato standard.")
    
    compagnia_file = st.selectbox("Seleziona la compagnia del file", ["MSC", "CMA", "HAPAG"])
    file_caricato = st.file_uploader("Scegli un file Excel o CSV", type=["xlsx", "xls", "csv"])
    
    if file_caricato is not None:
        if st.button("Elabora e Unisci Dati"):
            try:
                if file_caricato.name.endswith('.csv'):
                    df_nuovo = pd.read_csv(file_caricato)
                else:
                    df_nuovo = pd.read_excel(file_caricato)
                
                df_standard = pd.DataFrame()
                
                # Mappatura e conversione in stringa per la validità
                if compagnia_file == "MSC":
                    df_standard["POL"] = df_nuovo["Port_Loading"].astype(str).str.upper().str.strip()
                    df_standard["POD"] = df_nuovo["Port_Destination"].astype(str).str.upper().str.strip()
                    df_standard["Container"] = df_nuovo["Equipment_Type"].astype(str).str.upper().str.strip()
                    df_standard["Nolo"] = pd.to_numeric(df_nuovo["Ocean_Freight"], errors='coerce').fillna(0.0)
                    df_standard["Addizionali"] = pd.to_numeric(df_nuovo["Surcharges"], errors='coerce').fillna(0.0)
                    df_standard["Descrizione_Addizionali"] = df_nuovo["Surcharges_Description"].astype(str).fillna("")
                    df_standard["Spese_Imbarco"] = pd.to_numeric(df_nuovo["Local_Charges"], errors='coerce').fillna(0.0)
                    df_standard["Descrizione_Spese_Imbarco"] = df_nuovo["Local_Charges_Description"].astype(str).fillna("")
                    df_standard["Free_Time"] = df_nuovo["Free_Days"].astype(str)
                    df_standard["Validità"] = df_nuovo["Expiry_Date"].astype(str).fillna("")
                    df_standard["Note"] = df_nuovo["Remarks"].astype(str).fillna("")
                    
                elif compagnia_file == "CMA":
                    df_standard["POL"] = df_nuovo["POL_Code"].astype(str).str.upper().str.strip()
                    df_standard["POD"] = df_nuovo["POD_Code"].astype(str).str.upper().str.strip()
                    df_standard["Container"] = df_nuovo["Size_Type"].astype(str).str.upper().str.strip()
                    df_standard["Nolo"] = pd.to_numeric(df_nuovo["Base_Rate"], errors='coerce').fillna(0.0)
                    df_standard["Addizionali"] = pd.to_numeric(df_nuovo["Surcharges"], errors='coerce').fillna(0.0)
                    df_standard["Descrizione_Addizionali"] = df_nuovo["Surcharge_Details"].astype(str).fillna("")
                    df_standard["Spese_Imbarco"] = pd.to_numeric(df_nuovo["THC"], errors='coerce').fillna(0.0)
                    df_standard["Descrizione_Spese_Imbarco"] = df_nuovo["THC_Details"].astype(str).fillna("")
                    df_standard["Free_Time"] = df_nuovo["Demurrage_Free_Time"].astype(str)
                    df_standard["Validità"] = df_nuovo["Valid_To"].astype(str).fillna("")
                    df_standard["Note"] = df_nuovo["Notes"].astype(str).fillna("")
                    
                elif compagnia_file == "HAPAG":
                    df_standard["POL"] = df_nuovo["Origin"].astype(str).str.upper().str.strip()
                    df_standard["POD"] = df_nuovo["Destination"].astype(str).str.upper().str.strip()
                    df_standard["Container"] = df_nuovo["Type"].astype(str).str.upper().str.strip()
                    df_standard["Nolo"] = pd.to_numeric(df_nuovo["Freight"], errors='coerce').fillna(0.0)
                    df_standard["Addizionali"] = pd.to_numeric(df_nuovo["Add-ons"], errors='coerce').fillna(0.0)
                    df_standard["Descrizione_Addizionali"] = df_nuovo["Add-ons_Comments"].astype(str).fillna("")
                    df_standard["Spese_Imbarco"] = pd.to_numeric(df_nuovo["Origin_THC"], errors='coerce').fillna(0.0)
                    df_standard["Descrizione_Spese_Imbarco"] = df_nuovo["THC_Comments"].astype(str).fillna("")
                    df_standard["Free_Time"] = df_nuovo["Free_Time_Days"].astype(str)
                    df_standard["Validità"] = df_nuovo["Expiration"].astype(str).fillna("")
                    df_standard["Note"] = df_nuovo["General_Notes"].astype(str).fillna("")

                df_standard["Totale_Nolo"] = df_standard["Nolo"] + df_standard["Addizionali"]
                df_standard["Compagnia"] = compagnia_file
                df_standard["Origine"] = "Automatico"
                
                df_standard = df_standard.dropna(subset=["POL", "POD", "Container"])
                
                df_pulito_precedente = df_master[df_master["Compagnia"] != compagnia_file]
                df_finale = pd.concat([df_pulito_precedente, df_standard], ignore_index=True)
                
                salva_database(df_finale)
                st.success(f"File elaborato con successo per {compagnia_file}!")
                st.rerun()
                
            except Exception as e:
                st.error(f"Errore durante l'elaborazione del file Excel. Dettaglio: {e}")

# ==========================================
# TAB 3: INSERIMENTO MANUALE CON INTERVALLO DATA
# ==========================================
with tab_manuale:
    st.header("Inserisci una singola tariffa con intervallo di validità")
    with st.form("Form Inserimento Dettagliato"):
        col_m1, col_m2, col_m3 = st.columns(3)
        
        with col_m1:
            man_pol = st.text_input("Porto POL (Partenza)").upper().strip()
            man_carrier = st.text_input("Compagnia (es. MSC)").upper().strip()
            man_nolo = st.number_input("Nolo Base (€)", min_value=0.0, step=50.0)
            man_spese = st.number_input("Spese Imbarco (€)", min_value=0.0, step=10.0)
            man_desc_spese = st.text_input("Descrizione Spese Imbarco (es. THC + ISPS)")
            
        with col_m2:
            man_pod = st.text_input("Porto POD (Destinazione)").upper().strip()
            man_container = st.selectbox("Tipo Container", ["20FT", "40FT", "40HC"])
            man_addizionali = st.number_input("Addizionali (€)", min_value=0.0, step=10.0)
            man_desc_addizionali = st.text_input("Descrizione Add.li (es. BAF + GRI)")
            man_freetime = st.text_input("Free Time (es. 14 giorni det+dem)")
            
        with col_m3:
            # Sostituito il calendario con un campo di testo per l'intervallo libero
            man_validita = st.text_input("Data Validità (es. 01/05/2026-31/05/2026)")
            man_note = st.text_area("Campo Note Libero (es. valido solo per merce IMO)", height=150)
            
        invia_form = st.form_submit_button("Salva Tariffa nel Database")
        
        if invia_form:
            if man_pol and man_pod and man_carrier:
                totale_calcolato = man_nolo + man_addizionali
                
                nuova_riga = pd.DataFrame([{
                    "POL": man_pol, "POD": man_pod, "Compagnia": man_carrier,
                    "Container": man_container, "Nolo": man_nolo, 
                    "Addizionali": man_addizionali, "Descrizione_Addizionali": man_desc_addizionali,
                    "Totale_Nolo": totale_calcolato, "Spese_Imbarco": man_spese, 
                    "Descrizione_Spese_Imbarco": man_desc_spese, "Free_Time": man_freetime,
                    "Validità": man_validita, "Note": man_note, "Origine": "Manuale"
                }])
                
                df_aggiornato = pd.concat([df_master, nuova_riga], ignore_index=True)
                salva_database(df_aggiornato)
                st.success(f"Tariffa salvata con validità: {man_validita}!")
                st.rerun()
            else:
                st.error("I campi POL, POD e Compagnia sono obbligatori.")

# ==========================================
# TAB 4: ARCHIVIO E MANUTENZIONE DATABASE
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
