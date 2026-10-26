# -*- coding: utf-8 -*-
"""
Dashboard de Monitoreo de Modulaciones y Rechazos
Fuente: Google Sheets "SIC - Santa Fe 2026" (pestañas SIC y ACUMULADO)

Reglas de negocio:
  * Cada línea con datos (en SIC o ACUMULADO) es una MODULACIÓN.
  * Estatus entregado / completado / concluido  -> MODULACIÓN RESUELTA.
  * Rechazo válido = línea con motivo de rechazo (o estatus "rechazado")
    cuyo motivo NO sea "consulta ventas" ni "fletero informa".
  * No se calcula % de rechazo (solo conteos absolutos).

Ejecutar localmente:  streamlit run app.py
"""
from __future__ import annotations

import io
import re
import unicodedata
from urllib.parse import quote

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# =============================================================================
# 1. CONFIGURACIÓN GENERAL
# =============================================================================
st.set_page_config(
    page_title="Monitoreo de Modulaciones y Rechazos",
    page_icon="📦",
    layout="wide",
)

HOJA_SIC = "SIC"
HOJA_ACUM = "ACUMULADO"

# Motivos que NO deben contarse como rechazos operativos (se comparan normalizados)
MOTIVOS_EXCLUIDOS = ["consulta ventas", "fletero informa"]

# Palabras clave para detectar columnas (se busca "contiene", sin acentos/mayúsculas)
CLAVES_COLUMNAS = {
    "fecha": ["fecha", "dia", "date"],
    "chofer": ["chofer", "conductor"],
    "estatus": ["estatus", "estado", "status"],
    "motivo": ["motivo", "causa"],
}

# Estatus que cuentan como "resuelto" y los que lo niegan ("no entregado", etc.)
RE_RESUELTO = re.compile(r"\b(entreg|complet|conclu)")
RE_NEGADO = re.compile(r"\b(no|sin|pendiente de)\s+(entreg|complet|conclu)")
RE_RECHAZADO = re.compile(r"\brechaz")

ZONA_HORARIA = "America/Argentina/Buenos_Aires"


# =============================================================================
# 2. UTILIDADES
# =============================================================================
def normalizar(texto) -> str:
    """Minúsculas, sin acentos y sin espacios repetidos. Vacío si es nulo."""
    if texto is None or (not isinstance(texto, str) and pd.isna(texto)):
        return ""
    t = unicodedata.normalize("NFKD", str(texto))
    t = "".join(c for c in t if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", t).strip().lower()


MOTIVOS_EXCLUIDOS_NORM = {normalizar(m) for m in MOTIVOS_EXCLUIDOS}


def secreto(clave: str, defecto=None):
    """Lee st.secrets sin romper si no existe el archivo de secretos."""
    try:
        return st.secrets[clave]
    except Exception:
        return defecto


def extraer_id_planilla(texto: str) -> str:
    """Acepta el ID solo o la URL completa de la planilla."""
    m = re.search(r"/d/([a-zA-Z0-9_-]+)", texto or "")
    return m.group(1) if m else (texto or "").strip()


def hoy_local() -> pd.Timestamp:
    return pd.Timestamp.now(tz=ZONA_HORARIA).tz_localize(None).normalize()


# =============================================================================
# 3. CARGA DE DATOS (todo se lee "crudo", sin encabezado, y se procesa después)
# =============================================================================
@st.cache_data(ttl=300, show_spinner="Leyendo Google Sheets…")
def leer_google_sheets(sheet_id: str, usar_cuenta_servicio: bool) -> dict[str, pd.DataFrame]:
    """
    Lee las pestañas SIC y ACUMULADO.
      - Con cuenta de servicio (st.secrets['gcp_service_account']): funciona con planillas privadas.
      - Sin ella: usa el endpoint CSV público (la planilla debe estar compartida
        como "Cualquier persona con el enlace puede ver").
    """
    salida: dict[str, pd.DataFrame] = {}
    if usar_cuenta_servicio:
        import gspread
        from google.oauth2.service_account import Credentials

        creds = Credentials.from_service_account_info(
            dict(st.secrets["gcp_service_account"]),
            scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"],
        )
        libro = gspread.authorize(creds).open_by_key(sheet_id)
        for hoja in (HOJA_SIC, HOJA_ACUM):
            salida[hoja] = pd.DataFrame(libro.worksheet(hoja).get_all_values(), dtype=str)
    else:
        for hoja in (HOJA_SIC, HOJA_ACUM):
            url = (
                f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq"
                f"?tqx=out:csv&sheet={quote(hoja)}"
            )
            salida[hoja] = pd.read_csv(url, header=None, dtype=str, keep_default_na=False)
    return salida


@st.cache_data(show_spinner="Procesando archivo cargado…")
def leer_archivo_subido(contenido: bytes, nombre: str, origen_csv: str) -> dict[str, pd.DataFrame]:
    """Lee un Excel (busca pestañas SIC/ACUMULADO) o un CSV (se asigna al origen elegido)."""
    if nombre.lower().endswith(".csv"):
        df = pd.read_csv(io.BytesIO(contenido), header=None, dtype=str, keep_default_na=False)
        return {origen_csv: df}

    hojas = pd.read_excel(io.BytesIO(contenido), sheet_name=None, header=None, dtype=str)
    por_nombre = {h.strip().upper(): h for h in hojas}
    salida = {}
    for hoja in (HOJA_SIC, HOJA_ACUM):
        if hoja in por_nombre:
            salida[hoja] = hojas[por_nombre[hoja]].fillna("")
    return salida


def aplicar_encabezado(crudo: pd.DataFrame, fila_encabezado: int) -> pd.DataFrame:
    """Usa la fila indicada (1 = primera) como encabezado y limpia filas vacías."""
    if crudo.empty or len(crudo) < fila_encabezado:
        return pd.DataFrame()

    nombres, vistos = [], {}
    for i, c in enumerate(crudo.iloc[fila_encabezado - 1]):
        c = str(c).strip() or f"columna_{i + 1}"
        if c in vistos:                      # evita nombres de columna repetidos
            vistos[c] += 1
            c = f"{c}_{vistos[c]}"
        else:
            vistos[c] = 0
        nombres.append(c)

    df = crudo.iloc[fila_encabezado:].copy()
    df.columns = nombres
    df = df.replace(r"^\s*$", pd.NA, regex=True).dropna(how="all")
    return df.reset_index(drop=True)


def quitar_duplicados_sic(sic: pd.DataFrame, acum: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    """Si una fila de SIC ya fue volcada a ACUMULADO (idéntica en columnas comunes), se descarta de SIC."""
    comunes = [c for c in sic.columns if c in acum.columns]
    if not comunes or sic.empty or acum.empty:
        return sic, 0

    def clave(d: pd.DataFrame) -> pd.Series:
        return d[comunes].fillna("").astype(str).agg("|".join, axis=1)

    repetidas = clave(sic).isin(set(clave(acum)))
    return sic[~repetidas], int(repetidas.sum())


# =============================================================================
# 4. ESTANDARIZACIÓN Y REGLAS DE NEGOCIO
# =============================================================================
def detectar_columna(columnas: list[str], claves: list[str]) -> str | None:
    """Devuelve la primera columna cuyo nombre normalizado contiene alguna clave (por prioridad)."""
    for clave in claves:
        for col in columnas:
            if clave in normalizar(col):
                return col
    return None


def estandarizar(df: pd.DataFrame, origen: str, manual: dict[str, str], hoy: pd.Timestamp):
    """
    Convierte una pestaña al esquema estándar y calcula las banderas de negocio.
    Devuelve (DataFrame estandarizado, columnas detectadas).
    """
    cols = list(df.columns)
    detectadas = {}
    for rol, claves in CLAVES_COLUMNAS.items():
        elegida = manual.get(rol)
        detectadas[rol] = elegida if elegida in cols else detectar_columna(cols, claves)

    out = pd.DataFrame(index=df.index)
    out["origen"] = origen

    # --- Fecha -----------------------------------------------------------------
    if detectadas["fecha"]:
        out["fecha"] = pd.to_datetime(
            df[detectadas["fecha"]], errors="coerce", dayfirst=True, format="mixed"
        ).dt.normalize()
    else:
        out["fecha"] = pd.NaT
    if origen == HOJA_SIC:                    # SIC = día en curso: sin fecha válida -> hoy
        out["fecha"] = out["fecha"].fillna(hoy)

    # --- Chofer ----------------------------------------------------------------
    if detectadas["chofer"]:
        chofer = df[detectadas["chofer"]].fillna("").astype(str)
        chofer = chofer.str.replace(r"\s+", " ", regex=True).str.strip().str.upper()
    else:
        chofer = pd.Series("", index=df.index)
    out["chofer"] = chofer.replace("", "(SIN CHOFER)")

    # --- Estatus y motivo (texto original + versión normalizada) ---------------
    for rol in ("estatus", "motivo"):
        if detectadas[rol]:
            texto = df[detectadas[rol]].fillna("").astype(str).str.strip()
        else:
            texto = pd.Series("", index=df.index)
        out[rol] = texto
        out[f"{rol}_norm"] = texto.map(normalizar)

    # --- Banderas de negocio ---------------------------------------------------
    est = out["estatus_norm"]
    out["es_resuelta"] = est.str.contains(RE_RESUELTO) & ~est.str.contains(RE_NEGADO)

    rechazo_bruto = (out["motivo_norm"] != "") | est.str.contains(RE_RECHAZADO)
    out["es_excluido"] = out["motivo_norm"].isin(MOTIVOS_EXCLUIDOS_NORM)
    out["es_rechazo_valido"] = rechazo_bruto & ~out["es_excluido"]

    # --- Columnas originales restantes (para el detalle) -----------------------
    resto = df.drop(columns=[c for c in detectadas.values() if c], errors="ignore").copy()
    reservadas = set(out.columns)
    resto.columns = [f"{c} (hoja)" if c in reservadas else c for c in resto.columns]
    return pd.concat([out, resto], axis=1), detectadas


def asignar_periodo(fecha: pd.Series, granularidad: str) -> pd.Series:
    """Etiqueta de periodo ordenable: Día, Semana ISO o Mes."""
    if granularidad == "Día":
        return fecha.dt.strftime("%Y-%m-%d")
    if granularidad == "Semana":
        iso = fecha.dt.isocalendar()
        return iso["year"].astype(str) + "-S" + iso["week"].astype(str).str.zfill(2)
    return fecha.dt.strftime("%Y-%m")


def resumir(df: pd.DataFrame, por_chofer: bool) -> pd.DataFrame:
    """Tabla resumen por periodo (y chofer si corresponde)."""
    claves = ["periodo"] + (["chofer"] if por_chofer else [])
    resumen = (
        df.assign(modulacion=1)
        .groupby(claves, as_index=False)
        .agg(
            modulaciones=("modulacion", "sum"),
            resueltas=("es_resuelta", "sum"),
            rechazos=("es_rechazo_valido", "sum"),
            descartados=("es_excluido", "sum"),
        )
        .sort_values(claves)
    )
    return resumen.rename(columns={
        "periodo": "Periodo", "chofer": "Chofer", "modulaciones": "Modulaciones totales",
        "resueltas": "Modulaciones resueltas", "rechazos": "Rechazos válidos",
        "descartados": "Descartados (auditoría)",
    })


def para_mostrar(df: pd.DataFrame) -> pd.DataFrame:
    """Versión legible del detalle (nombres amigables, fechas sin hora, Sí/No)."""
    v = df.drop(columns=["estatus_norm", "motivo_norm"], errors="ignore").copy()
    v["fecha"] = v["fecha"].dt.date
    for b in ("es_resuelta", "es_rechazo_valido", "es_excluido"):
        v[b] = v[b].map({True: "Sí", False: "No"})
    return v.rename(columns={
        "origen": "Origen", "fecha": "Fecha", "periodo": "Periodo", "chofer": "Chofer",
        "estatus": "Estatus", "motivo": "Motivo", "es_resuelta": "Resuelta",
        "es_rechazo_valido": "Rechazo válido", "es_excluido": "Descartado",
    })


def exportar_excel(resumen: pd.DataFrame, detalle: pd.DataFrame, auditoria: pd.DataFrame) -> bytes:
    """Genera el .xlsx con tres hojas: Resumen, Detalle y Auditoría."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        resumen.to_excel(writer, sheet_name="Resumen", index=False)
        para_mostrar(detalle).to_excel(writer, sheet_name="Detalle", index=False)
        para_mostrar(auditoria).to_excel(writer, sheet_name="Auditoría", index=False)
        for hoja in writer.sheets.values():           # ancho de columnas legible
            for columna in hoja.columns:
                largo = max(len(str(c.value)) if c.value is not None else 0 for c in columna)
                hoja.column_dimensions[columna[0].column_letter].width = min(largo + 2, 45)
    return buffer.getvalue()


# =============================================================================
# 5. GRÁFICOS
# =============================================================================
def grafico_evolucion(resumen_periodo: pd.DataFrame, granularidad: str) -> go.Figure:
    """Barras agrupadas: Modulaciones Resueltas vs Rechazos Válidos por periodo."""
    fig = go.Figure()
    fig.add_bar(x=resumen_periodo["Periodo"], y=resumen_periodo["Modulaciones resueltas"],
                name="Modulaciones resueltas", marker_color="#2e7d32")
    fig.add_bar(x=resumen_periodo["Periodo"], y=resumen_periodo["Rechazos válidos"],
                name="Rechazos válidos", marker_color="#c62828")
    fig.update_layout(
        barmode="group", hovermode="x unified", height=420,
        title=f"Evolución por {granularidad.lower()}: resueltas vs rechazos válidos",
        xaxis=dict(title=granularidad, type="category"),
        yaxis=dict(title="Cantidad"),
        legend=dict(orientation="h", y=1.1),
        margin=dict(t=80, b=40),
    )
    return fig


def grafico_motivos(df: pd.DataFrame) -> go.Figure | None:
    """Top de motivos de rechazo válido (barras horizontales)."""
    rechazos = df[df["es_rechazo_valido"]]
    if rechazos.empty:
        return None
    conteo = (rechazos["motivo_norm"].replace("", "(sin motivo)")
              .value_counts().head(15).sort_values())
    fig = go.Figure(go.Bar(x=conteo.values, y=conteo.index, orientation="h",
                           marker_color="#c62828"))
    fig.update_layout(title="Motivos de rechazo válido", height=max(300, 28 * len(conteo) + 120),
                      xaxis_title="Cantidad", margin=dict(t=60, b=40))
    return fig


# =============================================================================
# 6. APLICACIÓN PRINCIPAL
# =============================================================================
def main():
    st.title("📦 Monitoreo de Modulaciones y Rechazos")
    st.caption("SIC - Santa Fe 2026 · Modulaciones del día (SIC) + histórico (ACUMULADO)")
    hoy = hoy_local()

    # ------------------------------ SIDEBAR: fuente -------------------------------
    sb = st.sidebar
    sb.header("🎛️ Controles")

    sb.subheader("Fuente de datos")
    archivo = sb.file_uploader("Cargar Excel/CSV manualmente (opcional)", type=["xlsx", "xls", "csv"])
    origen_csv = HOJA_ACUM
    if archivo is not None and archivo.name.lower().endswith(".csv"):
        origen_csv = sb.radio("El CSV corresponde a:", [HOJA_SIC, HOJA_ACUM], horizontal=True)

    if archivo is None:
        sheet_id = extraer_id_planilla(secreto("SHEET_ID", ""))
        if not sheet_id:
            sheet_id = extraer_id_planilla(sb.text_input("ID o URL de la planilla de Google"))
        if sb.button("🔄 Actualizar datos"):
            st.cache_data.clear()
            st.rerun()

    # ------------------------------ Carga -----------------------------------------
    try:
        if archivo is not None:
            crudos = leer_archivo_subido(archivo.getvalue(), archivo.name, origen_csv)
        elif sheet_id:
            crudos = leer_google_sheets(sheet_id, secreto("gcp_service_account") is not None)
        else:
            st.info("Configurá `SHEET_ID` en los secretos, pegá el ID de la planilla o subí un archivo.")
            st.stop()
    except Exception as e:  # noqa: BLE001
        st.error(f"No se pudo leer la fuente de datos: {e}")
        st.info("Verificá que la planilla esté compartida (o configurá la cuenta de servicio) "
                f"y que existan las pestañas «{HOJA_SIC}» y «{HOJA_ACUM}».")
        st.stop()

    if not crudos:
        st.warning(f"No se encontraron las pestañas «{HOJA_SIC}» / «{HOJA_ACUM}».")
        st.stop()

    # ------------------------------ SIDEBAR: configuración avanzada ---------------
    with sb.expander("⚙️ Configuración avanzada"):
        fila_enc = int(st.number_input("Fila del encabezado", min_value=1, value=1, step=1))
        quitar_dup = st.checkbox("Descartar de SIC filas ya presentes en ACUMULADO", value=True)
        frames = {h: aplicar_encabezado(c, fila_enc) for h, c in crudos.items()}
        opciones = ["(automático)"] + sorted({c for f in frames.values() for c in f.columns})
        manual = {}
        for rol, etiqueta in [("fecha", "Columna de fecha"), ("chofer", "Columna de chofer"),
                              ("estatus", "Columna de estatus"), ("motivo", "Columna de motivo")]:
            sel = st.selectbox(etiqueta, opciones, key=f"col_{rol}")
            if sel != "(automático)":
                manual[rol] = sel

    descartadas_dup = 0
    if quitar_dup and HOJA_SIC in frames and HOJA_ACUM in frames:
        frames[HOJA_SIC], descartadas_dup = quitar_duplicados_sic(frames[HOJA_SIC], frames[HOJA_ACUM])

    # ------------------------------ Estandarización -------------------------------
    partes, detectadas_por_hoja = [], {}
    for hoja, f in frames.items():
        if f.empty:
            continue
        std, detectadas = estandarizar(f, hoja, manual, hoy)
        partes.append(std)
        detectadas_por_hoja[hoja] = detectadas

    if not partes:
        st.warning("Las pestañas no tienen datos para analizar.")
        st.stop()

    base = pd.concat(partes, ignore_index=True)

    with sb.expander("🔎 Columnas detectadas"):
        for hoja, det in detectadas_por_hoja.items():
            st.write(f"**{hoja}**: " + " · ".join(f"{k} → `{v or '—'}`" for k, v in det.items()))
    for hoja, det in detectadas_por_hoja.items():
        faltan = [k for k in ("estatus", "motivo") if not det[k]]
        if faltan:
            st.warning(f"En «{hoja}» no se detectó la columna de {', '.join(faltan)}. "
                       "Seleccionala en *Configuración avanzada*.")

    sin_fecha = int(base["fecha"].isna().sum())
    base = base.dropna(subset=["fecha"])
    if sin_fecha:
        st.warning(f"{sin_fecha} fila(s) sin fecha válida fueron excluidas del análisis.")
    if base.empty:
        st.error("No hay filas con fecha válida. Revisá la columna de fecha.")
        st.stop()

    # ------------------------------ SIDEBAR: filtros ------------------------------
    sb.subheader("Filtros")
    granularidad = sb.radio("Temporalidad", ["Día", "Semana", "Mes"], horizontal=True,
                            help="Semana = semana ISO del año")
    f_min, f_max = base["fecha"].min().date(), base["fecha"].max().date()
    rango = sb.date_input("Rango de fechas", value=(f_min, f_max), min_value=f_min, max_value=f_max)
    if not (isinstance(rango, tuple) and len(rango) == 2):
        st.info("Seleccioná fecha de inicio y de fin.")
        st.stop()

    choferes = sorted(base["chofer"].unique())
    sel_choferes = sb.multiselect("Chofer(es)", choferes, placeholder="Todos")
    por_chofer = sb.checkbox("Desglosar tabla por chofer", value=bool(sel_choferes))

    # ------------------------------ Aplicar filtros -------------------------------
    df = base[(base["fecha"].dt.date >= rango[0]) & (base["fecha"].dt.date <= rango[1])]
    if sel_choferes:
        df = df[df["chofer"].isin(sel_choferes)]
    df = df.assign(periodo=asignar_periodo(df["fecha"], granularidad))

    if df.empty:
        st.warning("No hay datos con los filtros seleccionados.")
        st.stop()
    if descartadas_dup:
        st.caption(f"ℹ️ {descartadas_dup} fila(s) de SIC ya estaban en ACUMULADO y no se duplicaron.")

    # ------------------------------ KPIs ------------------------------------------
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Modulaciones totales", f"{len(df):,}".replace(",", "."))
    k2.metric("Modulaciones resueltas", f"{int(df['es_resuelta'].sum()):,}".replace(",", "."))
    k3.metric("Rechazos válidos", f"{int(df['es_rechazo_valido'].sum()):,}".replace(",", "."))
    k4.metric("Descartados (auditoría)", f"{int(df['es_excluido'].sum()):,}".replace(",", "."),
              help="Motivos «consulta ventas» y «fletero informa»: no cuentan como rechazo.")

    # ------------------------------ Pestañas --------------------------------------
    resumen = resumir(df, por_chofer)
    resumen_periodo = resumir(df, False)
    auditoria = df[df["es_excluido"]].sort_values("fecha")

    tab_resumen, tab_auditoria, tab_detalle = st.tabs(["📊 Resumen", "🔍 Auditoría", "📄 Detalle"])

    with tab_resumen:
        st.plotly_chart(grafico_evolucion(resumen_periodo, granularidad))
        st.subheader(f"Resumen por {granularidad.lower()}" + (" y chofer" if por_chofer else ""))
        st.dataframe(resumen, hide_index=True)

        fig_mot = grafico_motivos(df)
        if fig_mot:
            st.plotly_chart(fig_mot)

        st.download_button(
            "⬇️ Descargar resultado filtrado (.xlsx)",
            data=exportar_excel(resumen, df, auditoria),
            file_name=f"modulaciones_rechazos_{rango[0]}_{rango[1]}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    with tab_auditoria:
        st.subheader("Registros descartados del conteo de rechazos")
        st.caption("Motivos excluidos por regla de negocio: "
                   + ", ".join(f"«{m}»" for m in MOTIVOS_EXCLUIDOS)
                   + ". Se aplican los filtros de fecha y chofer.")
        if auditoria.empty:
            st.success("No hay registros descartados en el periodo seleccionado.")
        else:
            cols = st.columns(len(MOTIVOS_EXCLUIDOS))
            for col, motivo in zip(cols, MOTIVOS_EXCLUIDOS):
                col.metric(motivo.capitalize(),
                           int((auditoria["motivo_norm"] == normalizar(motivo)).sum()))
            st.dataframe(para_mostrar(auditoria), hide_index=True)

    with tab_detalle:
        st.subheader("Detalle de modulaciones filtradas")
        st.dataframe(para_mostrar(df.sort_values("fecha")), hide_index=True)


if __name__ == "__main__":
    main()
