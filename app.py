import streamlit as st
import pandas as pd
import os
import pdfplumber
import re

st.set_page_config(page_title="Tariffario Noli Marittimi", layout="wide", page_icon="🚢")

DB_FILE = "database_noli_periodi.csv"

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

# (Il TAB 1 rimane invariato per i colleghi)
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
        risultati = df_master[(df_master["POL"] == pol_scelto) & (df_master["POD"] == pod_scelto) & (df_master["Container"] == tipo_container)]
        if not risultati.empty:
            st.dataframe(risultati.sort_values(by="Totale_Nolo"), use_container_width=True)
        else:
            st.warning("Nessuna tariffa trovata.")

# ==========================================
# TAB 2: AGGIORNAMENTO AUTOMATICO (EXCEL E PDF)
# ==========================================
with tab_automatico:
    st.header("Caricamento Listini Excel o PDF")
    st.write("Il sistema supporta file Excel (.xlsx) e quotazioni formattate in PDF.")
    
    compagnia_file = st.selectbox("Seleziona la compagnia", ["MSC", "CMA", "HAPAG"])
    file_caricato = st.file_uploader("Trascina qui il file", type=["xlsx", "xls", "csv", "pdf"])
    
    if file_caricato is not None:
        # CASO 1: IL FILE È UN PDF
        if file_caricato.name.endswith('.pdf'):
            st.info("Rilevato file PDF. Estrazione testo in corso...")
            if st.button("Analizza e Importa PDF"):
                righe_estratte = []
                with pdfplumber.open(file_caricato) as pdf:
                    for pagina in pdf.pages:
                        testo = pagina.extract_text()
                        if testo:
                            for linea in testo.split("\n"):
                                # Cerca righe che contengono porti (parole in maiuscolo) e numeri (prezzi)
                                # Esempio semplicistico di ricerca pattern numerici nel PDF
                                cifre = re.findall(r'\b\d{3,4}\b', linea)
                                if len(cifre) >= 1:
                                    righe_estratte.append(linea)
                
                if righe_estratte:
                    st.write("Linee tariffarie individuate nel PDF (anteprima):")
                    st.code("\n".join(righe_estratte[:10]))
                    st.warning("I PDF non hanno colonne fisse. Per evitare inserimenti errati, copia le righe rilevate qui sopra e inseriscile nel tab 'Inserimento Manuale'.")
                else:
                    st.error("Impossibile trovare tabelle di prezzi leggibili in questo PDF. Potrebbe trattarsi di un PDF scannerizzato come immagine.")

        # CASO 2: IL FILE È UN ECCEL
        else:
            try:
                df_nuovo = pd.read_csv(file_caricato) if file_caricato.name.endswith('.csv') else pd.read_excel(file_caricato)
                st.success("File Excel caricato correttamente!")
                
                st.subheader("Mappatura colonne: Collega i campi del tuo Excel al database")
                st.write("Seleziona quale colonna del tuo file corrisponde ai nostri dati standard:")
                
                col_opzioni = list(df_nuovo.columns)
                
                c1, c2, c3 = st.columns(3)
                with c1:
                    m_pol = st.selectbox("Colonna Porto Partenza (POL)", col_opzioni)
                    m_nolo = st.selectbox("Colonna Prezzo Nolo", col_opzioni)
                with c2:
                    m_pod = st.selectbox("Colonna Porto Destinazione (POD)", col_opzioni)
                    m_add = st.selectbox("Colonna Prezzo Addizionali", col_opzioni)
                with c3:
                    m_cont = st.selectbox("Colonna Tipo Container", col_opzioni)
                    m_val = st.selectbox("Colonna Validità/Periodo", col_opzioni)
                
                if st.button("Conferma Mappatura e Unisci Excel"):
                    df_standard = pd.DataFrame()
                    df_standard["POL"] = df_nuovo[m_pol].astype(str).str.upper().str.strip()
                    df_standard["POD"] = df_nuovo[m_pod].astype(str).str.upper().str.strip()
                    df_standard["Container"] = df_nuovo[m_cont].astype(str).str.upper().str.strip()
                    df_standard["Nolo"] = pd.to_numeric(df_nuovo[m_nolo], errors='coerce').fillna(0.0)
                    df_standard["Addizionali"] = pd.to_numeric(df_nuovo[m_add], errors='coerce').fillna(0.0)
                    df_standard["Validità"] = df_nuovo[m_val].astype(str).fillna("")
                    
                    df_standard["Totale_Nolo"] = df_standard["Nolo"] + df_standard["Addizionali"]
                    df_standard["Compagnia"] = compagnia_file
                    df_standard["Origine"] = "Automatico"
                    df_standard["Descrizione_Addizionali"] = "Importato da Excel"
                    df_standard["Spese_Imbarco"] = 0.0
                    df_standard["Descrizione_Spese_Imbarco"] = ""
                    df_standard["Free_Time"] = ""
                    df_standard["Note"] = ""
                    
                    df_standard = df_standard.dropna(subset=["POL", "POD"])
                    df_pulito_precedente = df_master[df_master["Compagnia"] != compagnia_file]
                    df_finale = pd.concat([df_pulito_precedente, df_standard], ignore_index=True)
                    
                    salva_database(df_finale)
                    st.success("Dati Excel integrati con successo nel database condiviso!")
                    st.rerun()
            except Exception as e:
                st.error(f"Errore di lettura Excel: {e}")

# (I TAB 3 e 4 rimangono invariati rispetto a prima)
with tab_manuale:
    st.header("Inserisci una singola tariffa")
    # ... mantiene lo stesso codice precedente ...
with tab_database:
    st.header("Archivio")
    st.dataframe(df_master, use_container_width=True)
