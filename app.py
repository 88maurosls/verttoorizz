import streamlit as st
import pandas as pd
import numpy as np
import io
from PIL import Image
import datetime as dt

# =========================
# FIX: Size lette come date / seriali Excel
# =========================
EXCEL_EPOCH = pd.Timestamp("1899-12-30")   # base seriali Excel
SERIAL_OFFSET = 46141                     # 46150 -> 9  (size = seriale - 46141)
SIZE_MIN = 0.0
SIZE_MAX = 20.0

def excel_serial_from_datetime(x) -> int:
    ts = pd.Timestamp(x)
    return int((ts - EXCEL_EPOCH).days)

def normalize_size(val):
    """
    Ritorna size numerica (float) coerente.
    Gestisce:
    - datetime/date (lettura Excel come date) -> seriale -> size (seriale - offset)
    - seriali Excel numerici tipo 46150 -> size (x - offset)
    - numeri normali (6, 6.5)
    - stringhe ("6,5", " 7.5 ")
    """
    if pd.isna(val):
        return np.nan

    if isinstance(val, (pd.Timestamp, dt.datetime, dt.date)):
        serial = excel_serial_from_datetime(val)
        return float(serial - SERIAL_OFFSET)

    if isinstance(val, (int, float, np.integer, np.floating)):
        x = float(val)
        if x > 1000:
            return float(x - SERIAL_OFFSET)
        return x

    s = str(val).strip().replace(",", ".")
    if s.lower() in ("nan", "none", ""):
        return np.nan
    try:
        x = float(s)
        if x > 1000:
            return float(x - SERIAL_OFFSET)
        return x
    except:
        try:
            ts = pd.to_datetime(s, errors="raise")
            serial = excel_serial_from_datetime(ts)
            return float(serial - SERIAL_OFFSET)
        except:
            return np.nan

def nice_header(x: float):
    if abs(x - round(x)) < 1e-9:
        return int(round(x))
    return x

def to_wide(df: pd.DataFrame, sku_col: str, size_col: str, qty_col: str, add_tot: bool = True) -> pd.DataFrame:
    d = df.copy()

    # SKU
    d[sku_col] = d[sku_col].astype(str).str.strip()
    d = d[d[sku_col].notna() & (d[sku_col] != "")]

    # qty
    d[qty_col] = pd.to_numeric(d[qty_col], errors="coerce").fillna(0).astype(int)

    # size normalizzata
    d["_size_norm"] = d[size_col].apply(normalize_size)

    # range fisso 0-20
    d = d[d["_size_norm"].between(SIZE_MIN, SIZE_MAX, inclusive="both")]

    # pivot
    wide = (
        d.pivot_table(
            index=sku_col,
            columns="_size_norm",
            values=qty_col,
            aggfunc="sum",
            fill_value=0
        )
        .reset_index()
    )

    # ordina size
    size_cols = sorted([c for c in wide.columns if c != sku_col], key=float)
    wide = wide[[sku_col] + size_cols]

    # header puliti
    wide = wide.rename(columns={c: nice_header(float(c)) for c in size_cols})

    # TOT
    if add_tot:
        num_cols = [c for c in wide.columns if c != sku_col]
        wide.insert(1, "TOT", wide[num_cols].sum(axis=1).astype(int))

    return wide

# =========================
# Streamlit UI
# =========================
st.set_page_config(page_title="Da verticale a orizzontale (SKU x Size)", layout="wide")

st.title("Da verticale a orizzontale (SKU x Size)")
st.caption("Carica un Excel con colonne SKU, Size, qty. Output: SKU una sola volta, taglie in orizzontale, TOT.")

# Immagine esempio (facoltativa)
with st.expander("Mostra esempio (facoltativo)"):
    try:
        esempio_img = Image.open("eg.jpg")
        st.image(esempio_img, caption="Esempio", use_container_width=True)
    except FileNotFoundError:
        st.info("File 'eg.jpg' non trovato (puoi ignorare).")

file = st.file_uploader("Carica il file Excel", type=["xlsx"])

riga_header_excel = st.number_input(
    "Riga header (in Excel)",
    min_value=1,
    value=1,
    step=1,
    help="Esempio: se l'header è alla riga 7, inserisci 7."
)
header_idx = int(riga_header_excel) - 1

if file:
    try:
        df = pd.read_excel(file, engine="openpyxl", header=header_idx)

        # UI essenziale: scelta colonne
        st.subheader("Seleziona le colonne")
        cols = list(df.columns)

        def pick_default(options, candidates):
            lower_map = {str(o).strip().lower(): o for o in options}
            for c in candidates:
                if c in lower_map:
                    return lower_map[c]
            return options[0] if options else None

        sku_default = pick_default(cols, ["sku"])
        size_default = pick_default(cols, ["size", "taglia"])
        qty_default = pick_default(cols, ["qty", "qty ", "quantity", "quantità", "quantita"])

        c1, c2, c3 = st.columns(3)
        with c1:
            sku_col = st.selectbox("SKU", cols, index=cols.index(sku_default) if sku_default in cols else 0)
        with c2:
            size_col = st.selectbox("Size", cols, index=cols.index(size_default) if size_default in cols else 0)
        with c3:
            qty_col = st.selectbox("Qty", cols, index=cols.index(qty_default) if qty_default in cols else 0)

        add_tot = st.checkbox("Aggiungi colonna TOT", value=True)

        # Preview opzionale (pulita per l'operatore)
        with st.expander("Anteprima (facoltativa)"):
            st.dataframe(df, use_container_width=True)

        if st.button("Genera e scarica"):
            with st.spinner("Elaborazione in corso..."):
                src_total = int(pd.to_numeric(df[qty_col], errors="coerce").fillna(0).sum())

                out_df = to_wide(
                    df=df,
                    sku_col=sku_col,
                    size_col=size_col,
                    qty_col=qty_col,
                    add_tot=add_tot
                )

                out_total = int(out_df["TOT"].sum()) if "TOT" in out_df.columns else int(
                    out_df.drop(columns=[sku_col]).sum().sum()
                )

                # Messaggio semplice per l'operatore
                st.success(f"Fatto. Totale sorgente: {src_total} | Totale output: {out_total}")

                # Export
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    out_df.to_excel(writer, index=False, sheet_name="RESULT")
                output.seek(0)

                st.download_button(
                    label="📥 Scarica Excel",
                    data=output.getvalue(),
                    file_name="pivot_taglie.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"Errore: {e}")
