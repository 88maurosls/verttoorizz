import streamlit as st
import pandas as pd
import numpy as np
import io
from PIL import Image

# =========================
# Config parsing Size
# =========================
EXCEL_EPOCH = pd.Timestamp("1899-12-30")  # base per seriali Excel
SERIAL_OFFSET = 46141  # 46150 -> 9  (quindi size = seriale - 46141)

def normalize_size(val):
    """
    Ritorna una size numerica (float) in range 0-20.

    Gestisce:
    - numeri (6, 6.5)
    - stringhe ("6,5", " 7.5 ")
    - datetime (lette da Excel) convertendole a seriale e poi mappandole con offset
    - seriali Excel grandi (es. 46150) mappandoli con offset
    """
    if pd.isna(val):
        return np.nan

    # datetime (Excel letto come data)
    if isinstance(val, pd.Timestamp):
        serial = (val - EXCEL_EPOCH).days
        return float(serial - SERIAL_OFFSET)

    # numero
    if isinstance(val, (int, float, np.integer, np.floating)):
        x = float(val)
        if x > 1000:  # probabile seriale Excel
            return float(x - SERIAL_OFFSET)
        return x

    # stringa
    s = str(val).strip().replace(",", ".")
    if s.lower() in ("nan", "none", ""):
        return np.nan
    try:
        x = float(s)
        if x > 1000:
            return float(x - SERIAL_OFFSET)
        return x
    except:
        return np.nan

def nice_header(x: float):
    """6.0 -> 6, 6.5 resta 6.5"""
    if abs(x - round(x)) < 1e-9:
        return int(round(x))
    return x

def to_wide(df: pd.DataFrame, sku_col: str, size_col: str, qty_col: str,
            size_min: float = 0.0, size_max: float = 20.0, add_tot: bool = True) -> pd.DataFrame:
    d = df.copy()

    # Normalizza SKU
    d[sku_col] = d[sku_col].astype(str).str.strip()
    d = d[d[sku_col].notna() & (d[sku_col] != "")]

    # Qty
    d[qty_col] = pd.to_numeric(d[qty_col], errors="coerce").fillna(0).astype(int)

    # Size normalizzata
    d["_size_norm"] = d[size_col].apply(normalize_size)

    # Tieni solo range valido
    d = d[d["_size_norm"].between(size_min, size_max, inclusive="both")]

    # Pivot
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

    # Ordina colonne size
    size_cols = sorted([c for c in wide.columns if c != sku_col], key=float)
    wide = wide[[sku_col] + size_cols]

    # Intestazioni "belle"
    rename_map = {c: nice_header(float(c)) for c in size_cols}
    wide = wide.rename(columns=rename_map)

    # TOT
    if add_tot:
        num_cols = [c for c in wide.columns if c != sku_col]
        wide.insert(1, "TOT", wide[num_cols].sum(axis=1).astype(int))

    return wide

# =========================
# Streamlit UI
# =========================
st.set_page_config(page_title="Verticale → Orizzontale (SKU x Size)", layout="wide")

st.title("Verticale → Orizzontale (SKU x Size) v.2.0")
st.write(
    "Carica un file Excel in formato verticale (SKU, Size, qty) e genera un output "
    "con taglie in orizzontale, SKU una sola volta e TOT. "
    "Gestisce anche Size lette come date (seriali Excel tipo 46150 → 9)."
)

with st.expander("Esempio (opzionale)"):
    try:
        esempio_img = Image.open("eg.jpg")
        st.image(esempio_img, caption="Esempio", use_container_width=True)
    except FileNotFoundError:
        st.info("Immagine 'eg.jpg' non trovata (puoi ignorare).")

file = st.file_uploader("Carica il file Excel", type=["xlsx"])

riga_header_excel = st.number_input(
    "Riga dell'header Excel (es. 1)",
    min_value=1,
    value=1,
    step=1,
    help="Se i nomi colonna stanno in Excel alla riga 7, inserisci 7."
)
header_idx = int(riga_header_excel) - 1

if file:
    try:
        df = pd.read_excel(file, engine="openpyxl", header=header_idx)

        st.markdown("### Preview")
        st.dataframe(df, use_container_width=True)

        st.markdown("### Mappatura colonne")
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

        sku_col = st.selectbox("Colonna SKU", cols, index=cols.index(sku_default) if sku_default in cols else 0)
        size_col = st.selectbox("Colonna Size", cols, index=cols.index(size_default) if size_default in cols else 0)
        qty_col = st.selectbox("Colonna Quantità", cols, index=cols.index(qty_default) if qty_default in cols else 0)

        st.markdown("### Regole Size")
        c1, c2, c3 = st.columns(3)
        with c1:
            size_min = st.number_input("Size minima", value=0.0, step=0.5)
        with c2:
            size_max = st.number_input("Size massima", value=20.0, step=0.5)
        with c3:
            add_tot = st.checkbox("Aggiungi TOT", value=True)

        st.markdown("### Debug (utile se il totale non torna)")
        show_debug = st.checkbox("Mostra righe scartate e conteggi", value=True)

        if st.button("Genera file con taglie in orizzontale"):
            with st.spinner("Elaborazione in corso..."):
                # Totale sorgente (qty)
                src_total = pd.to_numeric(df[qty_col], errors="coerce").fillna(0).sum()

                # Normalizzazione size per debug
                tmp = df.copy()
                tmp["_size_norm"] = tmp[size_col].apply(normalize_size)
                tmp["_qty_num"] = pd.to_numeric(tmp[qty_col], errors="coerce").fillna(0).astype(int)

                kept = tmp[tmp["_size_norm"].between(size_min, size_max, inclusive="both")]
                dropped = tmp[~tmp["_size_norm"].between(size_min, size_max, inclusive="both")]

                out_df = to_wide(
                    df=df,
                    sku_col=sku_col,
                    size_col=size_col,
                    qty_col=qty_col,
                    size_min=size_min,
                    size_max=size_max,
                    add_tot=add_tot
                )

                out_total = out_df["TOT"].sum() if "TOT" in out_df.columns else out_df.drop(columns=[sku_col]).sum().sum()

                st.success("File generato!")

                st.markdown("### Totali")
                st.info(f"Totale qty sorgente: **{int(src_total)}** | Totale qty output: **{int(out_total)}**")

                st.markdown("### Anteprima output")
                st.dataframe(out_df, use_container_width=True)

                if show_debug:
                    st.markdown("### Debug: righe incluse / scartate")
                    st.write(f"Righe incluse: {len(kept)} | Somma qty incluse: {int(kept['_qty_num'].sum())}")
                    st.write(f"Righe scartate: {len(dropped)} | Somma qty scartate: {int(dropped['_qty_num'].sum())}")
                    st.markdown("#### Prime righe scartate (controllo)")
                    st.dataframe(dropped[[sku_col, size_col, "_size_norm", qty_col]].head(50), use_container_width=True)

                # Export excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    out_df.to_excel(writer, index=False, sheet_name="RESULT")
                output.seek(0)

                st.download_button(
                    label="📥 Scarica Excel (output)",
                    data=output.getvalue(),
                    file_name="pivot_taglie.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"❌ Errore: {e}")
