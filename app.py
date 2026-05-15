import streamlit as st
import pandas as pd
import os

# Configurazione iniziale dell'interfaccia web aziendale
st.set_page_config(page_title="Tariffario Noli Marittimi", layout="wide", page_icon="🚢")

DB_FILE = "database_noli_msc_specifico.csv"

def carica_database():
    if os.path.exists(DB_FILE):
        df = pd.read_csv(DB_FILE)
        df['Validità'] = df['Validità'].astype(str).fillna("")
        for col in ["Nolo", "Addizionali", "Totale_Nolo", "Spese_Imbarco", "BL"]:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)
        return df
    else:
        return pd.DataFrame(columns=[
            "POL", "POD", "Compagnia", "Container", 
            "Nolo", "Addizionali", "Descrizione_Addizionali", "Totale_Nolo", 
            "Spese_Imbarco", "Descrizione_Spese_Imbarco", "BL", "Free_Time", 
            "Validità", "Note", "Origine"
        ])

def salva_database(df):
    df.to_csv(DB_FILE, index=False)

df_master = carica_database()

st.title("🚢 Sistema Gestione Noli Marittimi - Configurazione MSC")
st.write("Calcolo automatico basato sulle specifiche MSC: Imbarco (THC, ISPS, LILO), Surcharges (EFS, BRC, ECA, ETS, FEU) e gestione della tassa BL.")

# Definizione dei 5 Tab ottimizzati
tab_ricerca, tab_automatico, tab_spese_porto, tab_manuale_singolo, tab_database = st.tabs([
    "🔍 Ricerca Tariffe", 
    "📂 1. Carica Matrice Excel", 
    "✍️ 2. Gestione Spese MSC per Porto",
    "➕ 3. Inserimento Manuale Spot",
    "📊 Archivio Database Completo"
])

# ==========================================
# TAB 1: INTERFACCIA DI RICERCA PER I COLLEGHI
# ==========================================
with tab_ricerca:
    st.header("Consultazione Tariffe e Dettaglio Costi")
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
            
            # Mostra anche la nuova colonna BL non sommata
            mostra_tabella = risultati[[
                "Compagnia", "Nolo", "Addizionali", "Descrizione_Addizionali", 
                "Totale_Nolo", "Spese_Imbarco", "Descrizione_Spese_Imbarco", 
                "BL", "Free_Time", "Validità", "Note"
            ]].copy()
            st.dataframe(
                mostra_tabella.style.format({
                    "Nolo": "€ {:.2f}", "Addizionali": "€ {:.2f}",
                    "Totale_Nolo": "€ {:.2f}", "Spese_Imbarco": "€ {:.2f}", "BL": "€ {:.2f}"
                }), use_container_width=True
            )
        else:
            st.warning("Nessuna tariffa corrispondente trovata.")

# ==========================================
# TAB 2: CARICAMENTO ESCLUSIVO DEI NOLI BASE
# ==========================================
with tab_automatico:
    st.header("Estrazione Automatica Noli Base da Matrice")
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
                    pol_raw = row.iloc
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
                            pod = pod.split("(").strip()
                            
                        for container_std in tipi_da_generare:
                            lista_tariffe.append({
                                "POL": pol, "POD": pod, "Compagnia": compagnia_file, "Container": container_std,
                                "Nolo": prezzo, "Addizionali": 0.0, "Descrizione_Addizionali": "", 
                                "Totale_Nolo": prezzo, "Spese_Imbarco": 0.0, "Descrizione_Spese_Imbarco": "", 
                                "BL": 0.0, "Free_Time": "", "Validità": validita_foglio, "Note": "Importato da matrice", "Origine": "Automatico"
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
# TAB 3: CONFIGURAZIONE SPECIFICA MSC PER PORTO (AGGIORNATO)
# ==========================================
with tab_spese_porto:
    st.header("✍️ Inserimento Spese Specifiche MSC per Porto")
    st.write("Scomponi i costi locali e le sottomisure tariffarie. I totali e le descrizioni si aggiorneranno da soli.")
    
    if not df_master.empty:
        lista_pol_esistenti = sorted(df_master["POL"].unique())
        pol_selezionato_spese = st.selectbox("Seleziona il Porto di Partenza (POL) da valorizzare", lista_pol_esistenti)
        container_selezionato_spese = st.selectbox("Applica al tipo container", ["20FT", "40FT", "40HC", "TUTTI"])
        
        st.markdown("---")
        col_an1, col_an2, col_an3 = st.columns(3)
        
        with col_an1:
            st.subheader("🏢 Spese Imbarco Locali")
            v_thc = st.number_input("THC (€/$)", min_value=0.0, step=5.0)
            v_isps = st.number_input("ISPS (€/$)", min_value=0.0, step=1.0)
            v_lilo = st.number_input("LILO (€/$)", min_value=0.0, step=5.0)
            
            totale_imb_calcolato = v_thc + v_isps + v_lilo
            st.metric("Totale Spese Imbarco", f"€ {totale_imb_calcolato:.2f}")
            
        with col_an2:
            st.subheader("📈 Surcharges / Addizionali")
            v_efs = st.number_input("EFS (€/$)", min_value=0.0, step=5.0)
            v_brc = st.number_input("BRC (€/$)", min_value=0.0, step=5.0)
            v_eca = st.number_input("ECA (€/$)", min_value=0.0, step=5.0)
            v_ets = st.number_input("ETS (€/$)", min_value=0.0, step=1.0)
            v_feu = st.number_input("FEU (€/$)", min_value=0.0, step=5.0)
            
            totale_add_calcolato = v_efs + v_brc + v_eca + v_ets + v_feu
            st.metric("Totale Addizionali", f"€ {totale_add_calcolato:.2f}")
            
        with col_an3:
            st.subheader("📄 Costo Spese Documentali")
            # CAMPO BL: Richiesto autonomo non sommato al totale nolo
            v_bl = st.number_input("Costo Documento BL (€/$)", min_value=0.0, step=5.0, help="Tassa fissa per polizza di carico, esclusa dal calcolo totale nolo.")
            st.write(" ")
            val_free_time = st.text_input("Free Time (Giorni det/dem)", "14 giorni free")
            val_note_libere = st.text_input("Note specifiche rotta", "Valido per nolo in vigore")
            
        if st.button(f"Calcola e Applica a tutte le rotte di {pol_selezionato_spese}"):
            df_modificato = df_master.copy()
            condizione = (df_modificato["POL"] == pol_selezionato_spese)
            if container_selezionato_spese != "TUTTI":
                condizione = condizione & (df_modificato["Container"] == container_selezionato_spese)
                
            if not df_modificato[condizione].empty:
                # Creazione automatica dei riepiloghi descrittivi basati sui nuovi campi
                testo_imb = f"THC:{v_thc} ISPS:{v_isps} LILO:{v_lilo}"
                
                lista_add_descr = []
                if v_efs > 0: lista_add_descr.append(f"EFS:{v_efs}")
                if v_brc > 0: lista_add_descr.append(f"BRC:{v_brc}")
                if v_eca > 0: lista_add_descr.append(f"ECA:{v_eca}")
                if v_ets > 0: lista_add_descr.append(f"ETS:{v_ets}")
                if v_feu > 0: lista_add_descr.append(f"FEU:{v_feu}")
                testo_add = " | ".join(lista_add_descr) if lista_add_descr else "Nessuna surcharge"
                
                # Salvataggio nel database centrale
                df_modificato.loc[condizione, "Spese_Imbarco"] = totale_imb_calcolato
                df_modificato.loc[condizione, "Descrizione_Spese_Imbarco"] = testo_imb
                df_modificato.loc[condizione, "Addizionali"] = totale_add_calcolato
                df_modificato.loc[condizione, "Descrizione_Addizionali"] = testo_add
                df_modificato.loc[condizione, "BL"] = v_bl  # Scrive il valore BL inserito
                df_modificato.loc[condizione, "Free_Time"] = val_free_time
                df_modificato.loc[condizione, "Note"] = val_note_libere
                
                # Calcolo finale: Nolo totale = Nolo base + Surcharges (il campo BL rimane escluso dal calcolo)
                df_modificato.loc[condizione, "Totale_Nolo"] = df_modificato.loc[condizione, "Nolo"] + totale_add_calcolato
                
                salva_database(df_modificato)
                st.success(f"Porto di {pol_selezionato_spese} configurato! Calcolo totali eseguito.")
                st.rerun()
    else:
        st.info("Nessun dato di nolo base presente. Esegui prima l'importazione nel Tab 1.")

# ==========================================
# TAB 4: INSERIMENTO MANUALE ANALITICO SPOT CON NUOVI CAMPI
# ==========================================
with tab_manuale_singolo:
    st.header("➕ Inserimento Manuale Singola Tariffa Spot Scomposta")
    with st.form("Form Inserimento Analitico"):
        col_m1, col_m2, col_m3 = st.columns(3)
        with col_m1:
            man_pol = st.text_input("Porto POL (Partenza)").upper().strip()
            man_carrier = st.text_input("Compagnia (es. MSC)").upper().strip()
            man_nolo = st.number_input("Nolo Base (€)", min_value=0.0, step=50.0)
            st.markdown("**Spese Locali Imbarco**")
            man_thc = st.number_input("THC (€)", min_value=0.0, step=5.0)
            man_isps = st.number_input("ISPS (€)", min_value=0.0, step=1.0)
            man_lilo = st.number_input("LILO (€)", min_value=0.0, step=5.0)
        with col_m2:
            man_pod = st.text_input("Porto POD (Destinazione)").upper().strip()
            man_container = st.selectbox("Tipo Container", ["20FT", "40FT", "40HC"])
            st.markdown("**Surcharges / Addizionali**")
            man_efs = st.number_input("EFS (€)", min_value=0.0, step=5.0)
            man_brc = st.number_input("BRC (€)", min_value=0.0, step=5.0)
            man_eca = st.number_input("ECA (€)", min_value=0.0, step=5.0)
            man_ets = st.number_input("ETS (€)", min_value=0.0, step=1.0)
            man_feu = st.number_input("FEU (€)", min_value=0.0, step=5.0)
        with col_m3:
            st.markdown("**Spese Documentali e Note**")
            man_bl = st.number_input("Costo Polizza BL (€)", min_value=0.0, step=5.0, help="Non viene inserito nel calcolo del totale nolo.")
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
                testo_imb_sin = f"THC:{man_thc} ISPS:{man_isps} LILO:{man_lilo}"
                
                nuova_riga = pd.DataFrame([{
                    "POL": man_pol, "POD": man_pod, "Compagnia": man_carrier, "Container": man_container,
                    "Nolo": man_nolo, "Addizionali": tot_add_sin, "Descrizione_Addizionali": testo_add_sin,
                    "Totale_Nolo": tot_nolo_sin, "Spese_Imbarco": tot_imb_sin, "Descrizione_Spese_Imbarco": testo_imb_sin,
                    "BL": man_bl, "Free_Time": man_freetime, "Validità": man_validita, "Note": man_note, "Origine": "Manuale"
                }])
                salva_database(pd.concat([df_master, nuova_riga], ignore_index=True))
                st.success("Tariffa spot salvata calcolando automaticamente i totali!")
                st.rerun()

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
