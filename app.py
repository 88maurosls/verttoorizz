import streamlit as st
import pandas as pd
import io
from PIL import Image

def parse_size_series(s: pd.Series) -> pd.Series:
    """
    Converte Size in numero quando possibile:
    - gestisce "6,5" -> 6.5
    - gestisce spazi
    - lascia NaN se non convertibile
    """
    s2 = s.astype(str).str.strip().str.replace(",", ".", regex=False)
    s2 = s2.replace({"nan": None, "None": None, "": None})
    return pd.to_numeric(s2, errors="coerce")

def to_wide(
    df: pd.DataFrame,
    sku_col: str,
    size_col: str,
    qty_col: str,
    size_min: float | None = None,
    size_max: float | None = None,
    keep_only_numeric_sizes: bool = True,
    add_tot: bool = True
) -> pd.DataFrame:
    # Copia di lavoro
    d = df.copy()

    # Colonne obbligatorie
    for c in (sku_col, size_col, qty_col):
        if c not in d.columns:
            raise ValueError(f"Colonna non trovata: {c}")

    # Normalizza SKU
    d[sku_col] = d[sku_col].astype(str).str.strip()
    d = d[d[sku_col].notna() & (d[sku_col] != "")]

    # Quantità
    d[qty_col] = pd.to_numeric(d[qty_col], errors="coerce").fillna(0).astype(int)

    # Size
    if keep_only_numeric_sizes:
        d["_size_num"] = parse_size_series(d[size_col])
        d = d[d["_size_num"].notna()]
        if size_min is not None:
            d = d[d["_size_num"] >= float(size_min)]
        if size_max is not None:
            d = d[d["_size_num"] <= float(size_max)]
        size_key = "_size_num"
    else:
        d["_size_txt"] = d[size_col].astype(str).str.strip()
        d = d[d["_size_txt"].notna() & (d["_size_txt"] != "")]
        size_key = "_size_txt"

    # Pivot: SKU una volta sola, taglie in orizzontale, somma qty
    wide = (
        d.pivot_table(
            index=sku_col,
            columns=size_key,
            values=qty_col,
            aggfunc="sum",
            fill_value=0
        )
        .reset_index()
    )

    # Ordina colonne taglia
    cols = list(wide.columns)
    base = [sku_col]
    size_cols = [c for c in cols if c not in base]

    if keep_only_numeric_sizes:
        size_cols_sorted = sorted(size_cols, key=lambda x: float(x))
        # intestazioni più belle: 6 invece di 6.0
        rename_map = {}
        for c in size_cols_sorted:
            v = float(c)
            if abs(v - round(v)) < 1e-9:
                rename_map[c] = int(round(v))
            else:
                rename_map[c] = v
        wide = wide.rename(columns=rename_map)
        # aggiorno lista size_cols dopo rename
        size_cols_sorted = [rename_map[c] for c in size_cols_sorted]
        ordered = base + size_cols_sorted
        wide = wide[ordered]
    else:
        size_cols_sorted = sorted(size_cols)
        wide = wide[base + size_cols_sorted]

    # Totale riga
    if add_tot:
        # tutte le colonne tranne SKU
        numeric_cols = [c for c in wide.columns if c != sku_col]
        wide.insert(1, "TOT", wide[numeric_cols].sum(axis=1))

    return wide


# ---------------- UI ----------------
st.title("Da verticale a orizzontale (SKU x Size) v.1.0")
st.write(
    "Carica un file Excel in formato verticale (es. colonne SKU, Size, qty) "
    "e genera un file con le taglie in orizzontale e lo SKU una sola volta."
)

# (Opzionale) immagine esempio
with st.expander("Mostra un esempio (opzionale)"):
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
        st.dataframe(df)

        st.markdown("### Mappatura colonne")
        cols = list(df.columns)

        # Scelte con default intelligenti
        def pick_default(options, candidates):
            lower_map = {str(o).strip().lower(): o for o in options}
            for c in candidates:
                if c in lower_map:
                    return lower_map[c]
            return options[0] if options else None

        sku_default = pick_default(cols, ["sku", "item code", "itemcode", "product", "codice"])
        size_default = pick_default(cols, ["size", "taglia"])
        qty_default = pick_default(cols, ["qty", "qty ", "quantity", "quantità", "quantita"])

        sku_col = st.selectbox("Colonna SKU", cols, index=cols.index(sku_default) if sku_default in cols else 0)
        size_col = st.selectbox("Colonna Size", cols, index=cols.index(size_default) if size_default in cols else 0)
        qty_col = st.selectbox("Colonna Quantità", cols, index=cols.index(qty_default) if qty_default in cols else 0)

        st.markdown("### Regole Size")
        keep_only_numeric = st.checkbox("Considera solo size numeriche", value=True)
        add_tot = st.checkbox("Aggiungi colonna TOT", value=True)

        c1, c2 = st.columns(2)
        with c1:
            size_min = st.number_input("Size minima (opzionale)", value=0.0, step=0.5) if keep_only_numeric else None
        with c2:
            size_max = st.number_input("Size massima (opzionale)", value=20.0, step=0.5) if keep_only_numeric else None

        if st.button("Genera file con taglie in orizzontale"):
            with st.spinner("Elaborazione in corso..."):
                out_df = to_wide(
                    df=df,
                    sku_col=sku_col,
                    size_col=size_col,
                    qty_col=qty_col,
                    size_min=size_min if keep_only_numeric else None,
                    size_max=size_max if keep_only_numeric else None,
                    keep_only_numeric_sizes=keep_only_numeric,
                    add_tot=add_tot
                )

                st.success("File generato!")
                st.markdown("### Anteprima output")
                st.dataframe(out_df)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    out_df.to_excel(writer, index=False, sheet_name="RESULT")
                output.seek(0)

                st.download_button(
                    label="Scarica Excel (output)",
                    data=output.getvalue(),
                    file_name="pivot_taglie.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"Errore durante la lettura o trasformazione: {e}")
