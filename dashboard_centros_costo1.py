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
    # Función auxiliar para normalizar texto
    def norm_txt(x):
        return (
            str(x)
            .strip()
            .lower()
            .replace("á","a").replace("é","e").replace("í","i").replace("ó","o").replace("ú","u")
        )

    # Intenta leer directamente; si falla, intenta con engine apropiado
    ext = str(getattr(file, "name", "")).lower()
    try:
        df = pd.read_excel(file)
    except Exception:
        try:
            file.seek(0)
            engine = "xlrd" if ext.endswith(".xls") else None
            df = pd.read_excel(file, engine=engine)
        except Exception as e:
            raise e

    # Si el archivo tiene filas de encabezado arriba (p.ej., títulos), detecta la fila header
    # Heurística: busca una fila con al menos 4 nombres de columnas esperadas
    expected_aliases = {
        "centro", "anio", "ano", "año", "mes", "cta", "descripcion", "descripción",
        "nit", "nombre", "valor", "ccunix", "rccunix", "codccp"
    }

    def has_headers(dcols):
        cols_n = [norm_txt(c) for c in dcols]
        hits = sum(1 for c in cols_n if c in expected_aliases)
        return hits >= 4

    # Si no detecta columnas clave, intenta redescubrir header en las primeras 20 filas
    if not has_headers(df.columns):
        file.seek(0)
        raw = pd.read_excel(file, header=None)
        header_row = None
        max_rows_check = min(25, len(raw))
        for i in range(max_rows_check):
            row_vals = [norm_txt(v) for v in raw.iloc[i].tolist()]
            hits = sum(1 for v in row_vals if v in expected_aliases)
            if hits >= 4:
                header_row = i
                break
        if header_row is not None:
            file.seek(0)
            df = pd.read_excel(file, skiprows=header_row, header=0)
        else:
            # Último intento: infiere con la primera fila no vacía como header
            file.seek(0)
            df = pd.read_excel(file, header=0)

    # Normaliza nombres de columnas con leniencia
    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.lower()
        .str.replace(r"[^a-z0-9]+", "_", regex=True)
    )

    # Renombrados lenientes
    colmap = {
        "anio": "anio",
        "ano": "anio",
        "año": "anio",
        "mes": "mes",
        "cta": "cta",
        "ccunix": "ccunix",
        "rccunix": "rccunix",
        "descripcion": "descripcion",
        "descripci_on": "descripcion",
        "nit": "nit",
        "nombre": "nombre",
        "valor": "valor",
        "centro": "centro",
        "codccp": "codccp",
    }
    # Mapeo flexible adicional
    inferred = {}
    for c in list(df.columns):
        if c in colmap:
            continue
        lc = c.lower()
        if "centro" in lc and "centro" not in df.columns:
            inferred[c] = "centro"
        elif "descrip" in lc and "descripcion" not in df.columns:
            inferred[c] = "descripcion"
        elif any(k in lc for k in ["valor", "monto", "importe"]) and "valor" not in df.columns:
            inferred[c] = "valor"
        elif "nombre" in lc and "nombre" not in df.columns:
            inferred[c] = "nombre"
        elif lc in ["a_o", "a\u00f1o"]:
            inferred[c] = "anio"
    colmap.update(inferred)
    df = df.rename(columns=colmap)

    # Filtra filas con columnas clave faltantes (post-limpieza)
    needed = ["centro","anio","mes","cta","descripcion","nit","nombre","valor"]
    missing = [c for c in needed if c not in df.columns]
    if missing:
        # Muestra columnas detectadas para debug amigable
        raise ValueError(f"Faltan columnas requeridas: {missing}. Columnas encontradas: {list(df.columns)}")

    # Limpieza tipos
    df["anio"] = pd.to_numeric(df["anio"], errors="coerce").astype("Int64")
    df["mes"] = pd.to_numeric(df["mes"], errors="coerce").astype("Int64")
    df["valor"] = pd.to_numeric(df["valor"], errors="coerce")

    # Remueve filas vacías
    df = df.dropna(subset=["anio","mes","valor"])

    # Crea fecha (usa día 1)
    df["fecha"] = pd.to_datetime(dict(year=df["anio"], month=df["mes"], day=1))

    # Normaliza textos
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
        var_pct = (monthly["valor"].iloc[-1] - monthly["valor"].iloc[-2]) / (monthly["valor"].iloc[-2] if monthly["valor"].iloc[-2] != 0 else np.nan)
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
        st.info("No hay datos para graficar.")
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
        m["var_pct"] = (m["valor"] - m["valor_prev"]) / m["valor_prev"]
        m["var_flag"] = np.where(m["var_pct"].abs() >= 0.25, "ALERTA", "")
    else:
        m["var_pct"] = np.nan
        st.caption("Si el error persiste, verifica que el archivo contenga columnas como: centro, año (anio/ano), mes, cta, descripcion, nit, nombre, valor. Puedes subir también la hoja exacta en .xlsx.")

        m["var_flag"] = ""
    st.subheader("Variación mensual")
    show = m.copy()
    show["Mes"] = pd.to_datetime(dict(year=show["anio"], month=show["mes"], day=1)).dt.strftime("%Y-%m")
    show = show[["Mes","valor","var_pct","var_flag"]]
    st.dataframe(
        show.style.format({"valor":"${:,.0f}","var_pct":"{:.1%}"}).apply(
            lambda s: ["background-color:#FDEDEC" if (isinstance(v, str) and v=="ALERTA") else "" for v in show["var_flag"]],
            axis=0
        ),
        use_container_width=True,
        height=300
    )

if uploaded:
    try:
        df = load_data(uploaded)
        # Filtros
        with st.sidebar:
            st.header("Filtros")
            centros = sorted(df["centro"].dropna().unique().tolist())
            ctas = sorted(df["cta"].dropna().unique().tolist())
            des = sorted(df["descripcion"].dropna().unique().tolist())
            provs = sorted(df["nombre"].dropna().unique().tolist())
            sel_centros = st.multiselect("Centro", centros, default=centros)
            sel_cta = st.multiselect("Cuenta (cta)", ctas, default=ctas)
            sel_desc = st.multiselect("Descripción", des, default=des)
            sel_prov = st.multiselect("Proveedor", provs, default=provs)
            miny, maxy = int(df["anio"].min()), int(df["anio"].max())
            yr = st.slider("Año", min_value=miny, max_value=maxy, value=(miny, maxy))
        mask = (
            df["centro"].isin(sel_centros) &
            df["cta"].isin(sel_cta) &
            df["descripcion"].isin(sel_desc) &
            df["nombre"].isin(sel_prov) &
            df["anio"].between(yr[0], yr[1])
        )
        dff = df[mask].copy()

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
            top3 = prov_conc.head(3).sum() / prov_conc.sum()
            notes.append(f"Concentración en top 3 proveedores: {top3:.1%}. Potencial de negociación y consolidación.")
        monthly = dff.groupby("fecha")["valor"].sum().sort_values()
        if len(monthly) >= 3:
            peak_m = monthly.idxmax().strftime("%Y-%m")
            notes.append(f"Mes pico de gasto: {peak_m}. Ver renovaciones/contratos.")
        if not notes:
            notes.append("La tendencia es estable sin eventos atípicos relevantes en el filtro actual.")
        for n in notes:
            st.write(f"• {n}")

    except Exception as e:
        st.error(f"Ocurrió un error leyendo el archivo: {e}")
else:
    st.info("Sube tu archivo para visualizar el dashboard.")
    st.image("https://placehold.co/1200x300?text=Sube+tu+Excel+para+ver+el+Dashboard")