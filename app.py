import streamlit as st
import pandas as pd
import io
from openpyxl.utils import column_index_from_string
from PIL import Image

# Funzione per trasporre taglie da un range di colonne
def trasponi_taglie(file, colonna_inizio, colonna_fine, riga_header):
    # Leggi il file Excel specificando la riga dell'header
    df = pd.read_excel(file, engine="openpyxl", header=riga_header)
    
    # Converti i riferimenti delle colonne (lettere) in indici numerici
    col_inizio_idx = column_index_from_string(colonna_inizio) - 1  # Indici 0-based
    col_fine_idx = column_index_from_string(colonna_fine)          # Indici 0-based + 1 per includere la colonna fine
    
    # Separa le colonne delle taglie e quelle rimanenti
    colonne_taglie = df.iloc[:, col_inizio_idx:col_fine_idx]  # Range di colonne per le taglie
    altre_colonne = df.iloc[:, :col_inizio_idx].join(df.iloc[:, col_fine_idx:])  # Colonne fuori dal range
    
    # Crea una lista per il dataframe trasposto
    righe = []
    for _, row in df.iterrows():
        for colonna in colonne_taglie.columns:
            if not pd.isna(row[colonna]):  # Salta celle vuote
                riga = row[:col_inizio_idx].to_dict()  # Copia valori prima del range
                riga.update(row[col_fine_idx:].to_dict())  # Copia valori dopo il range
                riga["Taglia"] = colonna  # Nome della colonna trasposta
                riga["Quantità"] = row[colonna]  # Valore corrispondente
                righe.append(riga)
    
    # Crea un nuovo dataframe
    df_trasposto = pd.DataFrame(righe)
    return df_trasposto

# Inizializza session_state
if 'file_elaborato' not in st.session_state:
    st.session_state.file_elaborato = None
if 'parametri_precedenti' not in st.session_state:
    st.session_state.parametri_precedenti = None

# Interfaccia Streamlit
st.title("Trasposizione di Colonne Taglie in Verticale v.1.2")
st.write("Carica il tuo file Excel e specifica il range di colonne da trasporre (es. taglie). Le altre colonne rimarranno invariate. NB: Eventuali filtri sul file di origine non verranno considerati e dovranno essere riapplicati sul nuovo file generato.")

# Visualizza un'immagine di esempio
st.markdown("### Esempio di input")
try:
    esempio_img = Image.open("eg.jpg")
    st.image(esempio_img, caption="Esempio di file Excel", use_container_width=True)
except FileNotFoundError:
    st.error("Immagine di esempio non trovata. Assicurati che 'eg.jpg' sia nella directory principale.")

# Caricamento del file Excel
file = st.file_uploader("Carica il file Excel", type=["xlsx"])

# Input per specificare la riga dell'header
riga_header_excel = st.number_input(
    "Riga dell'header Excel (es. 1)", 
    min_value=1, 
    value=1, 
    step=1,
    help="Specifica in quale riga si trovano i nomi delle colonne (es. se in Excel è riga 7, inserisci 7)"
)

# Converti in indice 0-based per pandas
riga_header = riga_header_excel - 1

# Input per specificare il range di colonne
colonna_inizio = st.text_input("Colonna inizio taglie (es. C)")
colonna_fine = st.text_input("Colonna fine taglie (es. Y)")

# Preview del range selezionato
if file and colonna_inizio and colonna_fine:
    try:
        st.markdown("### 📋 Preview del file")
        
        # Leggi il file con l'header specificato
        df_preview = pd.read_excel(file, engine="openpyxl", header=riga_header)
        
        # Converti i riferimenti delle colonne in indici
        col_inizio_idx = column_index_from_string(colonna_inizio) - 1
        col_fine_idx = column_index_from_string(colonna_fine)
        
        # Mostra informazioni sul range
        st.info(f"**Header dalla riga Excel:** {riga_header_excel} | **Range colonne:** {colonna_inizio} - {colonna_fine}")
        
        # Mostra tutto il dataframe
        st.write(f"**Anteprima file completo ({len(df_preview)} righe):**")
        st.dataframe(df_preview)
        
    except Exception as e:
        st.warning(f"Impossibile mostrare la preview: {e}")

# Elabora e scarica con spinner
if file and colonna_inizio and colonna_fine:
    try:
        # Crea una chiave univoca per i parametri correnti
        parametri_correnti = f"{file.name}_{riga_header}_{colonna_inizio}_{colonna_fine}"
        
        # Elabora solo se i parametri sono cambiati
        if st.session_state.parametri_precedenti != parametri_correnti:
            # Mostra spinner durante l'elaborazione
            with st.spinner('⏳ Elaborazione in corso... Sto trasponendo le taglie...'):
                # Trasforma il file e crea il nuovo dataframe
                nuovo_df = trasponi_taglie(file, colonna_inizio, colonna_fine, riga_header)
                
                # Salva in un file Excel temporaneo
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    nuovo_df.to_excel(writer, index=False, sheet_name="Trasposizione")
                output.seek(0)
                
                # Salva in session_state
                st.session_state.file_elaborato = output.getvalue()
                st.session_state.parametri_precedenti = parametri_correnti
        
        # Mostra il pulsante di download
        st.success("✅ File trasformato con successo!")
        st.download_button(
            label="📥 Scarica il file Excel trasformato",
            data=st.session_state.file_elaborato,
            file_name="trasposizione.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"❌ Errore durante la trasformazione: {e}")
