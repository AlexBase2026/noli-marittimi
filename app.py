import streamlit as st
import pandas as pd
import os

# Configurazione iniziale dell'interfaccia web aziendale
st.set_page_config(page_title="Tariffario Noli Marittimi", layout="wide", page_icon="🚢")

DB_FILE = "database_noli_ibrido.csv"

def carica_database():
    if os.path.exists(DB_FILE):
        df = pd.read_csv(DB_FILE)
        df['Validità'] = df['Validità'].astype(str).fillna("")
        # Assicura la corretta natura numerica dei campi di costo
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
st.title("🚢 Sistema Ibrido Gestione Noli Marittimi")
st.write("I noli base vengono estratti automaticamente dalle matrici Excel, mentre le spese locali e le addizionali si gestiscono manualmente per porto.")

tab_ricerca, tab_automatico, tab_spese_porto, tab_database = st.tabs([
    "🔍 Ricerca Tariffe", 
    "📂 1. Carica Matrice Noli Base", 
    "✍️ 2. Gestione Spese e Addizionali per Porto",
    "📊 Archivio Database Completo"
])

# ==========================================
# TAB 1: INTERFACCIA DI RICERCA
# ==========================================
with tab_ricerca:
    st.header("Consultazione Tariffe Aggiornate")
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
            
            # Formattazione e visualizzazione pulita
            mostra_tabella = risultati[[
                "Compagnia", "Nolo", "Addizionali", "Descrizione_Addizionali", 
                "Totale_Nolo", "Spese_Imbarco", "Descrizione_Spese_Imbarco", 
                "Free_Time", "Validità", "Note"
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
            st.warning("Nessuna tariffa corrispondente trovata.")

# ==========================================
# TAB 2: CARICAMENTO ESCLUSIVO DEI NOLI BASE
# ==========================================
with tab_automatico:
    st.header("Estrazione Automatica Noli Base")
    st.write("Carica la matrice Excel delle compagnie. Il sistema estrarrà i prezzi base sdoppiando i 40' in 40FT e 40HC.")
    
    compagnia_file = st.selectbox("Compagnia Marittima", ["MSC", "CMA", "HAPAG"])
    validita_foglio = st.text_input("Validità Temporale Foglio", "01/05/2026-31/05/2026")
    file_caricato = st.file_uploader("Trascina il file Excel (.xlsx)", type=["xlsx", "xls"])
    
    if file_caricato is not None:
        try:
            raw_df = pd.read_excel(file_caricato, header=None)
            st.success("File caricato. Pronto per l'estrazione.")
            
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
                                "Totale_Nolo": prezzo, # Inizialmente uguale al nolo base
                                "Spese_Imbarco": 0.0, "Descrizione_Spese_Imbarco": "", 
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
            st.error(f"Errore: {e}")

# ==========================================
# TAB 3: ASSEGNAZIONE MANUALE SPESE PER PORTO (LA TUA RICHIESTA)
# ==========================================
with tab_spese_porto:
    st.header("✍️ Gestione Spese e Addizionali Localizzate")
    st.write("Seleziona un porto di partenza inserito nel sistema per configurare o sovrascrivere le sue spese specifiche.")
    
    if not df_master.empty:
        lista_pol_esistenti = sorted(df_master["POL"].unique())
        pol_selezionato_spese = st.selectbox("Scegli il porto di partenza da aggiornare (POL)", lista_pol_esistenti)
        container_selezionato_spese = st.selectbox("Applica al tipo container", ["20FT", "40FT", "40HC", "TUTTI"])
        
        st.subheader(f"Inserisci le voci di costo per {pol_selezionato_spese}")
        
        col_s1, col_s2 = st.columns(2)
        with col_s1:
            val_spese_imb = st.number_input("Spese Imbarco (€/$) (es. THC + ISPS specifico)", min_value=0.0, step=5.0)
            desc_spese_imb = st.text_input("Descrizione Spese Imbarco", "THC + ISPS locale")
            val_free_time = st.text_input("Free Time dedicato", "14 giorni det/dem")
        with col_s2:
            val_addizionali = st.number_input("Addizionali Correttive (€/$)", min_value=0.0, step=5.0)
            desc_addizionali = st.text_input("Descrizione Addizionali", "EFS + BRC")
            val_note_libere = st.text_input("Note specifiche per questo porto", "Soggetto a variazioni locali")
            
        if st.button(f"Applica Costi a {pol_selezionato_spese}"):
            # Crea una copia di sicurezza del database
            df_modificato = df_master.copy()
            
            # Definisce la maschera di filtro per l'aggiornamento
            condizione = (df_modificato["POL"] == pol_selezionato_spese)
            if container_selezionato_spese != "TUTTI":
                condizione = condition = condizione & (df_modificato["Container"] == container_selezionato_spese)
                
            # Verifica se ci sono righe da aggiornare
            if df_modificato[condizione].empty:
                st.warning("Nessuna rotta trovata nel database per i filtri selezionati.")
            else:
                # Aggiorna i campi manuali nelle righe corrispondenti
                df_modificato.loc[condizione, "Spese_Imbarco"] = val_spese_imb
                df_modificato.loc[condizione, "Descrizione_Spese_Imbarco"] = desc_spese_imb
                df_modificato.loc[condizione, "Addizionali"] = val_addizionali
                df_modificato.loc[condizione, "Descrizione_Addizionali"] = desc_addizionali
                df_modificato.loc[condizione, "Free_Time"] = val_free_time
                df_modificato.loc[condizione, "Note"] = val_note_libere
                
                # Ricalcola il Totale Nolo (Nolo Base + Nuove Addizionali)
                df_modificato.loc[condizione, "Totale_Nolo"] = df_modificato.loc[condizione, "Nolo"] + val_addizionali
                
                salva_database(df_modificato)
                st.success(f"Database aggiornato! Le spese per il porto di {pol_selezionato_spese} sono state inserite e sommate a tutti i rispettivi noli base.")
                st.rerun()
    else:
        st.info("Il database è vuoto. Carica prima una matrice di noli base nel Tab 1.")

# ==========================================
# TAB 4: VISUALIZZAZIONE DATABASE
# ==========================================
with tab_database:
    st.header("Visualizzazione Tabellare di Controllo")
    st.dataframe(df_master, use_container_width=True)
    if st.button("🗑️ Svuota Database"):
        if os.path.exists(DB_FILE):
            os.remove(DB_FILE)
        st.rerun()
