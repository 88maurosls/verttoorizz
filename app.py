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

def excel_serial_from_datetime(x) -> int:
    # accetta datetime/date/pd.Timestamp
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

    # pandas Timestamp OR python datetime/date
    if isinstance(val, (pd.Timestamp, dt.datetime, dt.date)):
        serial = excel_serial_from_datetime(val)
        return float(serial - SERIAL_OFFSET)

    # numeri
    if isinstance(val, (int, float, np.integer, np.floating)):
        x = float(val)
        if x > 1000:  # probabile seriale Excel
            return float(x - SERIAL_OFFSET)
        return x

    # stringhe
    s = str(val).strip().replace(",", ".")
    if s.lower() in ("nan", "none", ""):
        return np.nan
    try:
        x = float(s)
        if x > 1000:
            return float(x - SERIAL_OFFSET)
        return x
    except:
        # ultima chance: stringa che sembra data
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

def to_wide(df: pd.DataFrame, sku_col: str, size_col: str, qty_col: str,
            size_min: float = 0.0, size_max: float = 20.0, add_tot: bool = True) -> pd.DataFrame:
    d = df.copy()

    # SKU
    d[sku_col] = d[sku_col].astype(str).str.strip()
    d = d[d[sku_col].notna() & (d[sku_col] != "")]

    # qty
    d[qty_col] = pd.to_numeric(d[qty_col], errors="coerce").fillna(0).astype(int)

    # size normalizzata
    d["_size_norm"] = d[size_col].apply(normalize_size)

    # filtro range
    d = d[d["_size_norm"].between(size_min, size_max, inclusive="both")]

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

    # ordina size numericamente
    size_cols = sorted([c for c in wide.columns if c != sku_col], key=float)
    wide = wide[[sku_col] + size_cols]

    # header più puliti
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

st.title("Verticale → Orizzontale (SKU x Size) v.2.1")
st.write(
    "Converte un file verticale (SKU, Size, qty) in un file con taglie in orizzontale.\n"
    "Fix incluso: se alcune Size vengono lette come date (es. 2026-05-11), vengono rimappate correttamente."
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
    step=1
)
header_idx = int(riga_header_excel) - 1

if file:
    try:
        # Leggo “grezzo” (non forzo parse date, ma Excel può contenere proprio date reali)
        df = pd.read_excel(file, engine="openpyxl", header=header_idx)

        st.markdown("### Preview (grezza)")
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

        show_debug = st.checkbox("Mostra debug normalizzazione Size", value=True)

        if show_debug:
            st.markdown("### Debug preview (Size normalizzata)")
            dbg = df[[sku_col, size_col, qty_col]].copy()
            dbg["_size_norm"] = dbg[size_col].apply(normalize_size)
            dbg["_qty_num"] = pd.to_numeric(dbg[qty_col], errors="coerce").fillna(0).astype(int)
            dbg["_kept_0_20"] = dbg["_size_norm"].between(size_min, size_max, inclusive="both")
            st.dataframe(dbg.tail(40), use_container_width=True)

            kept_sum = int(dbg.loc[dbg["_kept_0_20"], "_qty_num"].sum())
            dropped_sum = int(dbg.loc[~dbg["_kept_0_20"], "_qty_num"].sum())
            src_sum = int(dbg["_qty_num"].sum())
            st.info(f"Somma qty sorgente: {src_sum} | incluse (0-20): {kept_sum} | scartate: {dropped_sum}")

        if st.button("Genera file con taglie in orizzontale"):
            with st.spinner("Elaborazione in corso..."):
                # Totale sorgente
                src_total = int(pd.to_numeric(df[qty_col], errors="coerce").fillna(0).sum())

                out_df = to_wide(
                    df=df,
                    sku_col=sku_col,
                    size_col=size_col,
                    qty_col=qty_col,
                    size_min=size_min,
                    size_max=size_max,
                    add_tot=add_tot
                )

                out_total = int(out_df["TOT"].sum()) if "TOT" in out_df.columns else int(out_df.drop(columns=[sku_col]).sum().sum())

                st.success("File generato!")
                st.info(f"Totale qty sorgente: **{src_total}** | Totale qty output: **{out_total}**")

                st.markdown("### Anteprima output")
                st.dataframe(out_df, use_container_width=True)

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
