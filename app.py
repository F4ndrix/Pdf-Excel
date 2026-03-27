import streamlit as st
import pandas as pd
import fitz  # pymupdf
import re
import io

# Configurazione interfaccia per il Deploy
st.set_page_config(page_title="PDF → Excel", page_icon="📦", layout="wide")

# CSS personalizzato per stabilizzare la UI e migliorare l'aspetto
st.markdown("""
    <style>
    .stAlert { margin-top: 1rem; }
    .stButton>button { 
        width: 100%; 
        border-radius: 8px; 
        font-weight: bold; 
        background-color: #FF4B4B; 
        color: white;
        height: 3em;
    }
    .stDataFrame { border: 1px solid #e6e9ef; border-radius: 8px; }
    /* Nasconde messaggi di errore superflui della libreria */
    .stException { display: none; }
    </style>
    """, unsafe_allow_html=True)

st.title("📦 PDF → Magazzino Excel")
st.write("Strumento ottimizzato per bolle IDG.")

# Uso di colonne per organizzare meglio lo spazio
col1, col2 = st.columns([2, 1])

with col1:
    files = st.file_uploader("Trascina qui le tue Bolle PDF", type="pdf", accept_multiple_files=True, key="file_uploader_main")

with col2:
    st.info("💡 Consiglio: Se l'app dà errore grafico, ricarica la pagina del browser.")
    filtro_attivo = st.checkbox("🔍 Filtro Anti-Spazzatura", value=True, key="filtro_checkbox")

st.markdown("---")

# Struttura colonne ufficiale richiesta
colonne_magazzino = [
    "FAM", "CODICE", "DESCRIZIONE", "QUANTITA", "UDM", "PREZZO", 
    "SC1", "SC2", "SC3", "K", "L", "M", "PREZZO NETTO", "COMMESSA", 
    "QT PRE", "AC", "N BOLLA", "DATA", "FORNITORE", 
    "TOTALE UTILIZZAT", "TOTALE ACQUISTA", "TOT. BOLLA"
]

def clean_num(val):
    """Converte i formati numerici dei PDF (es. 1.250,00) in numeri elaborabili."""
    if not val: return 0.0
    try:
        # Rimuove simboli valuta e spazi, gestisce il formato italiano
        s = str(val).replace("€", "").replace(" ", "").strip()
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")
        elif "," in s:
            s = s.replace(",", ".")
        return float(s)
    except:
        return 0.0

def is_valid_number(token):
    """Verifica se una parola nel PDF è un numero (Qta o Prezzo)."""
    clean = token.replace(".", "").replace(",", "").replace("-", "").replace("€", "").strip()
    return clean.isdigit()

if files:
    righe_estratte = []
    
    for file in files:
        commessa_corrente = ""
        fornitore_rilevato = ""
        data_rilevata = ""
        
        try:
            file_bytes = file.read()
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            testo_intero = "".join([p.get_text() for p in doc])
            
            # --- IDENTIFICAZIONE FORNITORE ---
            lower_text = testo_intero.lower()
            if "sacchi" in lower_text: fornitore_rilevato = "SACCHI"
            elif "idg" in lower_text: fornitore_rilevato = "IDG"
            elif "doppler" in lower_text: fornitore_rilevato = "DOPPLER"
            
            # --- DATA DOCUMENTO ---
            match_data = re.search(r"(\d{2}[/-]\d{2}[/-]\d{2,4})", testo_intero)
            if match_data: data_rilevata = match_data.group(1)

            for pagina in doc:
                words = pagina.get_text("words")
                words.sort(key=lambda w: (w[1], w[0]))
                
                linee_geometriche = []
                linea_corrente = []
                y_corrente = None
                
                for w in words:
                    if y_corrente is None or abs(w[1] - y_corrente) <= 5:
                        linea_corrente.append(w)
                        y_corrente = w[1]
                    else:
                        linea_corrente.sort(key=lambda x: x[0])
                        linee_geometriche.append(" ".join([x[4] for x in linea_corrente]))
                        linea_corrente = [w]
                        y_corrente = w[1]
                if linea_corrente:
                    linea_corrente.sort(key=lambda x: x[0])
                    linee_geometriche.append(" ".join([x[4] for x in linea_corrente]))

                for linea in linee_geometriche:
                    linea = linea.strip()
                    if not linea: continue
                    l_low = linea.lower()

                    # GESTIONE COMMESSA / ORDINE
                    if any(x in l_low for x in ["ordine", "vs. ord", "commessa", "riferimento"]):
                        m = re.search(r"(?:ordine|ord\.|commessa|riferimento)\s*[:\s]*([\w/-]+)", linea, re.IGNORECASE)
                        if m: 
                            val = m.group(1).strip()
                            if len(val) > 2: commessa_corrente = val
                        continue

                    # FILTRI RIMOZIONE RIGA
                    if filtro_attivo:
                        if any(x in l_low for x in ["pag.1", "telefono", "sede legale", "p.iva", "www.", "pec:", "cap.soc", "destinatario", "spett.le", "banca appoggio"]): continue
                        if "disp." in l_low: continue 

                    tokens = linea.split()
                    nums = []
                    udm = ""
                    idx_end = len(tokens)
                    
                    # Analisi da destra verso sinistra
                    for i in range(len(tokens)-1, -1, -1):
                        t = tokens[i]
                        # Doppler a volte usa P2 invece di PZ
                        t_upper = t.upper().replace("P2", "PZ")
                        if is_valid_number(t):
                            nums.insert(0, t)
                            idx_end = i
                        elif t_upper in ["PZ", "NR", "MT", "RT", "KG", "CAD"]:
                            udm = t_upper
                            idx_end = i
                        elif t in ["%", "€"]:
                            idx_end = i
                        else: break
                    
                    if len(nums) < 2: continue
                    
                    desc_part = tokens[:idx_end]
                    if not desc_part or len(desc_part) < 1: continue

                    row = {c: "" for c in colonne_magazzino}
                    row["COMMESSA"] = commessa_corrente
                    row["FORNITORE"] = fornitore_rilevato
                    row["DATA"] = data_rilevata
                    row["UDM"] = udm
                    row["N BOLLA"] = "" # Sempre vuota come richiesto

                    # LOGICA SUDDIVISIONE FAM / CODICE / DESCRIZIONE
                    if len(desc_part) >= 2:
                        # Se il primo pezzo è molto corto o maiuscolo di 3 lettere, è probabilmente la FAM
                        if len(desc_part[0]) <= 4:
                            row["FAM"] = desc_part[0]
                            row["CODICE"] = desc_part[1]
                            row["DESCRIZIONE"] = " ".join(desc_part[2:])
                        else:
                            row["CODICE"] = desc_part[0]
                            row["DESCRIZIONE"] = " ".join(desc_part[1:])
                    else:
                        row["DESCRIZIONE"] = " ".join(desc_part)

                    qta = clean_num(nums[0])
                    prezzo_unitario = clean_num(nums[1])
                    
                    row["QUANTITA"] = qta
                    row["PREZZO"] = prezzo_unitario
                    row["PREZZO NETTO"] = round(prezzo_unitario, 4)
                    row["TOTALE ACQUISTA"] = round(prezzo_unitario * qta, 2)
                    
                    # Evita righe spazzatura se la descrizione è un numero
                    if not row["CODICE"] and not row["FAM"] and len(row["DESCRIZIONE"]) < 3:
                        continue
                        
                    righe_estratte.append(row)
            
            doc.close()

        except Exception as e:
            st.error(f"Errore file {file.name}: {e}")

    if righe_estratte:
        df = pd.DataFrame(righe_estratte, columns=colonne_magazzino)
        st.success(f"✅ Trovate {len(righe_estratte)} righe.")
        
        # Visualizzazione tabella con chiave unica per evitare removeChild error
        st.dataframe(df, use_container_width=True, key="data_editor_output")

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Dati_Magazzino')
        
        st.download_button(
            label="📥 SCARICA EXCEL",
            data=buf.getvalue(),
            file_name="Export_Magazzino.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            key="download_btn_final"
        )