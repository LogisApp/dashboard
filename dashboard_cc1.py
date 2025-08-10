import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Centro de Costos Unix - O28", page_icon="📊", layout="wide")

st.title("📊 Dashboard Centros de Costo Unix - O28")
st.caption("Sube tu archivo Excel con columnas: centro, anio, mes, cta, CCUNIX, rccunix, descripcion, nit, NOMBRE, valor, codccp.")

uploaded = st.file_uploader("Sube el archivo Excel (.xls o .xlsx)", type=["xls", "xlsx"])

@st.cache_data
def load_data(file):
    # Normalizador de texto
    def norm_txt(x: str) -> str:
        return (
            str(x)
            .strip()
            .lower()
            .replace("á","a").replace("é","e").replace("í","i").replace("ó","o").replace("ú","u")
        )

    def normalize_cols(df0: pd.DataFrame) -> pd.DataFrame:
        df0 = df0.copy()
        df0.columns = (
            df0.columns.astype(str)
            .str.strip()
            .str.lower()
            .str.replace(r"[^a-z0-9]+", "_", regex=True)
        )
        return df0

    def map_columns(df0: pd.DataFrame) -> pd.DataFrame:
        colmap = {
            "anio": "anio",
            "ano": "anio",
            "año": "anio",
            "mes": "mes",
            "cta": "cta",
            "cuenta": "cta",
            "cuenta_contable": "cta",
            "ccunix": "ccunix",
            "rccunix": "rccunix",
            "descripcion": "descripcion",
            "descripci_on": "descripcion",
            "nit": "nit",
            "nombre": "nombre",
            "proveedor": "nombre",
            "valor": "valor",
            "vlr": "valor",
            "total": "valor",
            "centro": "centro",
            "centro_de_costo": "centro",
            "codccp": "codccp",
        }
        inferred = {}
        for c in list(df0.columns):
            if c in colmap:
                continue
            lc = c.lower()
            if "centro" in lc and "centro" not in df0.columns:
                inferred[c] = "centro"
            elif "descrip" in lc and "descripcion" not in df0.columns:
                inferred[c] = "descripcion"
            elif any(k in lc for k in ["valor","monto","importe","total","vlr"]) and "valor" not in df0.columns:
                inferred[c] = "valor"
            elif any(k in lc for k in ["nombre","proveedor"]) and "nombre" not in df0.columns:
                inferred[c] = "nombre"
            elif lc in ["a_o", "a\u00f1o"]:
                inferred[c] = "anio"
        if inferred:
            colmap.update(inferred)
        return df0.rename(columns=colmap)

    def try_header_discovery(df_like: pd.DataFrame, file_like) -> pd.DataFrame:
        expected_aliases = {
            "centro", "anio", "ano", "año", "mes", "cta", "cuenta", "cuenta contable", "descripcion", "descripción",
            "nit", "nombre", "proveedor", "valor", "total", "vlr", "ccunix", "rccunix", "codccp"
        }
        # ¿Ya parece tener headers?
        cols_n = [norm_txt(c) for c in df_like.columns]
        if sum(1 for c in cols_n if c in expected_aliases) >= 4:
            return df_like
        # Buscar fila de headers
        file_like.seek(0)
        raw = pd.read_excel(file_like, header=None)
        header_row = None
        for i in range(min(50, len(raw))):
            row_vals = [norm_txt(v) for v in raw.iloc[i].tolist()]
            if sum(1 for v in row_vals if v in expected_aliases) >= 4:
                header_row = i
                break
        if header_row is not None:
            file_like.seek(0)
            return pd.read_excel(file_like, skiprows=header_row, header=0)
        # Último intento: usar primera fila como header
        file_like.seek(0)
        return pd.read_excel(file_like, header=0)

    # 1) Leer primer intento
    ext = str(getattr(file, "name", "")).lower()
    try:
        df = pd.read_excel(file)
    except Exception:
        file.seek(0)
        engine = "xlrd" if ext.endswith(".xls") else None
        df = pd.read_excel(file, engine=engine)

    # 2) Intentar descubrir encabezados
    df = try_header_discovery(df, file)

    # 3) Normalizar columnas y mapear
    df = normalize_cols(df)
    df = map_columns(df)

    # 4) Validación; si falla, intentar buscar en otras hojas
    needed = ["centro","anio","mes","cta","descripcion","nit","nombre","valor"]
    missing = [c for c in needed if c not in df.columns]
    if missing:
        # Buscar en todas las hojas
        file.seek(0)
        all_sheets = pd.read_excel(file, sheet_name=None)
        found = None
        for sname, sdf in all_sheets.items():
            # Para cada hoja, intentar descubrir header y normalizar
            tmpf = file  # reutiliza handler
            # Para re-leer correctamente cada hoja, cargamos de nuevo desde dict
            df_try = sdf.copy()
            df_try = normalize_cols(df_try)
            df_try = map_columns(df_try)
            miss2 = [c for c in needed if c not in df_try.columns]
            if len(miss2) == 0:
                found = df_try
                break
        if found is not None:
            df = found
        else:
            raise ValueError(f"Faltan columnas requeridas: {missing}. Columnas encontradas: {list(df.columns)}")

    # 5) Tipos
    df["anio"] = pd.to_numeric(df["anio"], errors="coerce").astype("Int64")
    df["mes"] = pd.to_numeric(df["mes"], errors="coerce").astype("Int64")
    df["valor"] = pd.to_numeric(df["valor"], errors="coerce")

    # 6) Limpieza
    df = df.dropna(subset=["anio","mes","valor"])

    # 7) Fecha
    df["fecha"] = pd.to_datetime(dict(year=df["anio"], month=df["mes"], day=1))

    # 8) Normaliza textos
    for tcol in ["centro","cta","descripcion","nombre","nit","ccunix","rccunix","codccp"]:
        if tcol in df.columns:
            df[tcol] = df[tcol].astype(str).str.strip()

    return df

def kpi_block(df):
    total = df["valor"].sum()
    # Variación vs mes anterior (sobre total)
    monthly = df.groupby("fecha", as_index=False)["valor"].sum().sort_values("fecha")
    var_pct = np.nan
    if len(monthly) >= 2:
        last_value = monthly["valor"].iloc[-1]
        prev_value = monthly["valor"].iloc[-2]
        if prev_value != 0:
            var_pct = (last_value - prev_value) / prev_value
            
    ticket = df.groupby(["nit","nombre"])["valor"].mean().mean() if not df.empty else np.nan

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Gasto total", f"${total:,.0f}")
    c2.metric("Variación mensual", f"{var_pct*100:,.1f}%" if pd.notna(var_pct) else "N/A")
    c3.metric("Ticket promedio proveedor", f"${ticket:,.0f}" if pd.notna(ticket) else "N/A")
    # Proveedor top
    prov = df.groupby(["nit","nombre"])["valor"].sum().sort_values(ascending=False).head(1)
    if len(prov) == 1:
        (nit, nom), val = prov.index[0], prov.iloc[0]
        c4.metric("Proveedor Top", f"{nom}", delta=f"${val:,.0f}")
    else:
        c4.metric("Proveedor Top", "N/A")

def trend_chart(df):
    m = df.groupby("fecha", as_index=False)["valor"].sum().sort_values("fecha")
    if m.empty:
        st.info("No hay datos para graficar con los filtros seleccionados.")
        return
    # Promedio móvil 3 meses
    m["prom_mov_3m"] = m["valor"].rolling(window=3, min_periods=1).mean()

    fig = go.Figure()
    fig.add_trace(go.Bar(x=m["fecha"], y=m["valor"], name="Gasto mensual", marker_color="#2E86DE"))
    fig.add_trace(go.Scatter(x=m["fecha"], y=m["prom_mov_3m"], name="Tendencia (PM 3M)", mode="lines+markers", line=dict(color="#E67E22", width=3)))
    fig.update_layout(
        title="Gasto mensual con línea de tendencia",
        xaxis_title="Mes",
        yaxis_title="COP",
        template="simple_white",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        bargap=0.2,
        height=420,
    )
    fig.update_yaxes(tickprefix="$", separatethousands=True)
    st.plotly_chart(fig, use_container_width=True)

def top_tables(df):
    left, right = st.columns(2)
    top_prov = df.groupby(["nit","nombre"], as_index=False)["valor"].sum().sort_values("valor", ascending=False).head(10)
    top_centros = df.groupby(["centro"], as_index=False)["valor"].sum().sort_values("valor", ascending=False).head(10)
    with left:
        st.subheader("Top 10 Proveedores")
        st.dataframe(top_prov.style.format({"valor":"${:,.0f}"}), use_container_width=True, height=300)
    with right:
        st.subheader("Top 10 Centros")
        st.dataframe(top_centros.style.format({"valor":"${:,.0f}"}), use_container_width=True, height=300)

def variance_table(df):
    m = df.groupby(["anio","mes"], as_index=False)["valor"].sum().sort_values(["anio","mes"])
    if len(m) >= 2:
        m["valor_prev"] = m["valor"].shift(1)
        m["var_pct"] = (m["valor"] - m["valor_prev"]) / m["valor_prev"].replace(0, np.nan)
        m["var_flag"] = np.where(m["var_pct"].abs() >= 0.25, "ALERTA", "")
    else:
        m["var_pct"] = np.nan
        m["var_flag"] = ""
    st.subheader("Variación mensual")
    show = m.copy()
    show["Mes"] = pd.to_datetime(dict(year=show["anio"], month=show["mes"], day=1)).dt.strftime("%Y-%m")
    show = show[["Mes","valor","var_pct","var_flag"]]
    st.dataframe(
        show.style.format({"valor":"${:,.0f}","var_pct":"{:.1%}"}).apply(
            lambda s: ["background-color:#FDEDEC" if (isinstance(v, str) and v=="ALERTA") else "" for v in s.values],
            subset=["var_flag"]
        ),
        use_container_width=True,
        height=300
    )

if uploaded:
    try:
        df = load_data(uploaded)
        
        # --- Callback para sincronizar la lógica de "Todos" en los filtros ---
        def sync_filters(filter_key):
            # Obtener el estado actual y el anterior
            current_selection = st.session_state[filter_key]
            prev_selection_key = f"prev_{filter_key}"
            previous_selection = st.session_state.get(prev_selection_key, ["Todos"])

            # 1. Si el usuario deselecciona todo, volver a "Todos"
            if not current_selection:
                st.session_state[filter_key] = ["Todos"]
            # 2. Si el usuario agrega "Todos" a una selección existente de items específicos
            elif "Todos" in current_selection and "Todos" not in previous_selection:
                st.session_state[filter_key] = ["Todos"]
            # 3. Si el usuario agrega un item específico mientras "Todos" está seleccionado
            elif "Todos" in previous_selection and len(current_selection) > 1:
                st.session_state[filter_key] = [item for item in current_selection if item != "Todos"]

            # Actualizar el estado anterior para la próxima interacción
            st.session_state[prev_selection_key] = st.session_state[filter_key]

        # --- Inicializar Session State para filtros si no existen ---
        filter_keys = ['sel_centros', 'sel_cta', 'sel_desc', 'sel_prov']
        for key in filter_keys:
            if key not in st.session_state:
                st.session_state[key] = ['Todos']
            prev_key = f"prev_{key}"
            if prev_key not in st.session_state:
                st.session_state[prev_key] = ['Todos']
        
        # --- Filtros en la barra lateral ---
        with st.sidebar:
            st.header("Filtros")
            
            # Crear listas de opciones con "Todos" al principio
            centros = ["Todos"] + sorted(df["centro"].dropna().unique().tolist())
            ctas = ["Todos"] + sorted(df["cta"].dropna().unique().tolist())
            des = ["Todos"] + sorted(df["descripcion"].dropna().unique().tolist())
            provs = ["Todos"] + sorted(df["nombre"].dropna().unique().tolist())

            # Widgets de selección múltiple con lógica de "Todos"
            st.multiselect("Centro", centros, key="sel_centros", on_change=sync_filters, args=("sel_centros",))
            st.multiselect("Cuenta (cta)", ctas, key="sel_cta", on_change=sync_filters, args=("sel_cta",))
            st.multiselect("Descripción", des, key="sel_desc", on_change=sync_filters, args=("sel_desc",))
            st.multiselect("Proveedor", provs, key="sel_prov", on_change=sync_filters, args=("sel_prov",))
            
            # --- Filtro de año (CORREGIDO) ---
            miny, maxy = int(df["anio"].min()), int(df["anio"].max())
            
            if miny == maxy:
                st.write(f"Mostrando datos para el único año disponible: **{miny}**")
                yr = (miny, maxy) # Establece el rango al único año
            else:
                yr = st.slider("Año", min_value=miny, max_value=maxy, value=(miny, maxy))

        # --- Lógica de filtrado ---
        dff = df.copy()
        if "Todos" not in st.session_state.sel_centros:
            dff = dff[dff["centro"].isin(st.session_state.sel_centros)]
        if "Todos" not in st.session_state.sel_cta:
            dff = dff[dff["cta"].isin(st.session_state.sel_cta)]
        if "Todos" not in st.session_state.sel_desc:
            dff = dff[dff["descripcion"].isin(st.session_state.sel_desc)]
        if "Todos" not in st.session_state.sel_prov:
            dff = dff[dff["nombre"].isin(st.session_state.sel_prov)]
            
        # Siempre se aplica el filtro de año
        dff = dff[dff["anio"].between(yr[0], yr[1])]
        
        # --- Renderizado del Dashboard ---
        if dff.empty:
            st.warning("No se encontraron datos con los filtros seleccionados.")
        else:
            kpi_block(dff)
            trend_chart(dff)
            top_tables(dff)
            variance_table(dff)

            # Insights automáticos
            st.subheader("Insights")
            notes = []
            if (dff["valor"] < 0).any():
                neg_sum = dff.loc[dff["valor"]<0,"valor"].sum()
                notes.append(f"Se detectan ajustes/valores negativos por ${neg_sum:,.0f}. Revisar notas contables.")
            
            prov_conc = dff.groupby(["nit","nombre"])["valor"].sum().sort_values(ascending=False)
            if len(prov_conc) >= 3 and prov_conc.sum() > 0:
                top3_pct = prov_conc.head(3).sum() / prov_conc.sum()
                notes.append(f"Concentración en top 3 proveedores: {top3_pct:.1%}. Potencial de negociación y consolidación.")
            
            monthly = dff.groupby("fecha")["valor"].sum().sort_values()
            if len(monthly) >= 3:
                peak_m = monthly.idxmax().strftime("%Y-%m")
                notes.append(f"Mes pico de gasto: {peak_m}. Ver renovaciones/contratos.")
            
            if not notes:
                notes.append("La tendencia es estable sin eventos atípicos relevantes en el filtro actual.")
            
            for n in notes:
                st.write(f"• {n}")

    except Exception as e:
        st.error(f"Ocurrió un error procesando el archivo: {e}")
        st.exception(e) # Muestra el traceback para facilitar la depuración
else:
    st.info("Sube tu archivo para visualizar el dashboard.")
    st.image("https://placehold.co/1200x300?text=Sube+tu+Excel+para+ver+el+Dashboard")
