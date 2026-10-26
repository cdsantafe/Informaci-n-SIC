"""Dashboard de Monitoreo de Modulaciones y Rechazos - SIC Santa Fe 2026.

- Lee solo el Google Sheets (sin carga manual).
- Se actualiza cada 15 min entre 6:30 y 15:00, de lunes a sábado.
- Por defecto muestra el día actual y la semana en curso.
- Rechazo = línea con datos cuyo Status es "Rechazo".
"""
import io
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import pandas as pd
import plotly.express as px
import streamlit as st
from streamlit_autorefresh import st_autorefresh

# ------------------------- Configuración -------------------------
SHEET_ID = "1X76M-VVd3CwbwMSTshxKubXrLljCrGf3Ejy-kC5ztuo"
GID_SIC = "1763302350"      # pestaña SIC (día en curso)
GID_ACUMULADO = ""          # pegar acá el gid de la pestaña ACUMULADO (histórico)
TZ = ZoneInfo("America/Argentina/Buenos_Aires")
INTERVALO_MIN = 15
INICIO, FIN = (6, 30), (15, 0)  # ventana de actualización (lunes a sábado)
RESUELTOS = ("entregado", "completado", "concluido")
FILTROS = ["Chofer", "Supervisor", "Vendedor"]  # se usan los que existan en la hoja

st.set_page_config(page_title="Monitoreo de Modulaciones y Rechazos", layout="wide")


# ------------------------- Horario / caché -------------------------
def ventana_activa(t):
    """True de lunes a sábado entre 6:30 y 15:00."""
    return t.weekday() != 6 and INICIO <= (t.hour, t.minute) < FIN


def clave_cache(t):
    """Cambia cada 15 min dentro de la ventana; fuera de ella, una vez por tramo."""
    if ventana_activa(t):
        return f"{t:%Y-%m-%d}-{t.hour:02d}{t.minute // INTERVALO_MIN * INTERVALO_MIN:02d}"
    tramo = "pre" if (t.hour, t.minute) < INICIO else "post"
    return f"fuera-{t:%Y-%m-%d}-{tramo}"


# ------------------------- Carga y limpieza -------------------------
def leer_hoja(gid):
    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={gid}"
    crudo = pd.read_csv(url, header=None, dtype=str, keep_default_na=False)
    for i, fila in crudo.iterrows():  # buscar la fila de encabezados
        vals = [v.strip().lower() for v in fila]
        if "status" in vals and "planilla" in vals:
            df = crudo.iloc[i + 1:].copy()
            df.columns = [v.strip() for v in fila]
            return df.loc[:, df.columns != ""].reset_index(drop=True)
    raise ValueError("No se encontró la fila de encabezados (Planilla / Status).")


def preparar(df):
    d = df.copy()
    d["Fecha"] = pd.to_datetime(d["Dia-Hora"], dayfirst=True, errors="coerce")
    estado = d["Status"].str.strip().str.lower()
    d["Tipo"] = "En curso / otro"
    d.loc[estado.isin(RESUELTOS), "Tipo"] = "Resuelta"
    d.loc[estado == "rechazo", "Tipo"] = "Rechazo"  # única regla de rechazo
    d["Bultos_n"] = pd.to_numeric(
        d["Bultos"].str.replace(".", "", regex=False).str.replace(",", ".", regex=False),
        errors="coerce",
    ).fillna(0)
    valido = d["Fecha"].notna() & (estado != "")
    descartados = d[~valido].replace("", pd.NA).dropna(how="all").fillna("")
    validos = d[valido].drop_duplicates("Planilla", keep="first")  # SIC pisa a ACUMULADO
    return validos, descartados


@st.cache_data(show_spinner="Cargando datos del Google Sheets…", max_entries=4)
def cargar_todo(clave):
    partes = [leer_hoja(GID_SIC)] + ([leer_hoja(GID_ACUMULADO)] if GID_ACUMULADO else [])
    validos, descartados = preparar(pd.concat(partes, ignore_index=True))
    return validos, descartados, datetime.now(TZ)


# ------------------------- Cálculos y gráficos -------------------------
def kpis(d):
    r = d["Tipo"]
    return {
        "Modulaciones totales": len(d),
        "Resueltas": int((r == "Resuelta").sum()),
        "Rechazos": int((r == "Rechazo").sum()),
        "Bultos rechazados": int(d.loc[r == "Rechazo", "Bultos_n"].sum()),
    }


def tarjetas(titulo, d):
    st.subheader(titulo)
    for col, (k, v) in zip(st.columns(4), kpis(d).items()):
        col.metric(k, f"{v:,}".replace(",", "."))


def con_periodo(d, modo):
    f = d["Fecha"]
    clave = {"Hora": f.dt.strftime("%H:00"), "Día": f.dt.strftime("%Y-%m-%d"),
             "Semana ISO": f.dt.strftime("%G-S%V"), "Mes": f.dt.strftime("%Y-%m")}[modo]
    return d.assign(Periodo=clave)


def grafico(d, modo):
    x = con_periodo(d, modo)
    g = x[x["Tipo"].isin(["Resuelta", "Rechazo"])].groupby(["Periodo", "Tipo"]).size()
    g = g.reset_index(name="Cantidad")
    fig = px.bar(g, x="Periodo", y="Cantidad", color="Tipo", barmode="group",
                 color_discrete_map={"Resuelta": "#2e7d32", "Rechazo": "#c62828"})
    st.plotly_chart(fig)


def resumen(d, modo, por):
    x = con_periodo(d, modo)
    x["Resueltas"] = x["Tipo"] == "Resuelta"
    x["Rechazos"] = x["Tipo"] == "Rechazo"
    x["Bultos_rechazados"] = x["Bultos_n"].where(x["Rechazos"], 0)
    claves = ["Periodo"] + ([por] if por else [])
    return x.groupby(claves).agg(Modulaciones=("Tipo", "size"), Resueltas=("Resueltas", "sum"),
                                 Rechazos=("Rechazos", "sum"),
                                 Bultos_rechazados=("Bultos_rechazados", "sum")).reset_index()


def a_excel(resumen_df, detalle, descartados):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        resumen_df.to_excel(w, "Resumen", index=False)
        detalle.to_excel(w, "Detalle", index=False)
        descartados.to_excel(w, "Descartados", index=False)
    return buf.getvalue()


# ------------------------- App -------------------------
t = datetime.now(TZ)
if t.weekday() != 6 and (t.hour, t.minute) < FIN:
    st_autorefresh(interval=INTERVALO_MIN * 60_000, key="auto")  # recarga la página

st.title("Dashboard de Monitoreo de Modulaciones y Rechazos")

try:
    datos, descartados, actualizado = cargar_todo(clave_cache(t))
except Exception as e:
    st.error(f"No se pudo leer el Google Sheets: {e}")
    st.info("Verificá que la hoja esté compartida como 'Cualquier persona con el enlace'.")
    st.stop()

st.caption(f"Última actualización: {actualizado:%d/%m/%Y %H:%M:%S} · se actualiza cada "
           f"{INTERVALO_MIN} min de 6:30 a 15:00, lunes a sábado")
if not GID_ACUMULADO:
    st.warning("Falta configurar GID_ACUMULADO: la semana solo incluye lo que haya en la pestaña SIC.")

# Barra lateral
vista = st.sidebar.radio("Vista", ["Hoy y semana en curso", "Solo hoy",
                                   "Solo semana en curso", "Personalizado"])
hoy = t.date()
lunes = hoy - timedelta(days=hoy.weekday())
for col in FILTROS:
    if col in datos.columns:
        sel = st.sidebar.multiselect(col, sorted(v for v in datos[col].unique() if v.strip()))
        if sel:
            datos = datos[datos[col].isin(sel)]
por = next((c for c in FILTROS if c in datos.columns), None)

dia = datos["Fecha"].dt.date
d_hoy = datos[dia == hoy]
d_sem = datos[(dia >= lunes) & (dia <= hoy)]

modo = "Día"
if vista == "Solo hoy":
    foco, modo = d_hoy, "Hora"
elif vista == "Personalizado":
    rango = st.sidebar.date_input("Rango de fechas", (lunes, hoy))
    modo = st.sidebar.radio("Agrupar por", ["Día", "Semana ISO", "Mes"])
    if len(rango) != 2:
        st.info("Elegí fecha de inicio y de fin.")
        st.stop()
    foco = datos[(dia >= rango[0]) & (dia <= rango[1])]
else:
    foco = d_sem

tab1, tab2, tab3 = st.tabs(["📊 Resumen", "📋 Detalle", "🔍 Auditoría"])

with tab1:
    if vista in ("Hoy y semana en curso", "Solo hoy"):
        tarjetas(f"Hoy ({hoy:%d/%m/%Y})", d_hoy)
    if vista in ("Hoy y semana en curso", "Solo semana en curso"):
        tarjetas(f"Semana en curso ({lunes:%d/%m} al {hoy:%d/%m})", d_sem)
    if vista == "Personalizado":
        tarjetas("Rango seleccionado", foco)
    if foco.empty:
        st.info("No hay modulaciones para el período elegido.")
    else:
        grafico(foco, modo)

res = resumen(foco, modo, por) if not foco.empty else pd.DataFrame()
cols = [c for c in ["Planilla", "Fecha", "Cliente", "Razon Social", "Supervisor", "Vendedor",
                    "Motivo", "Status", "Bultos_n", "Tipo"] if c in foco.columns]
with tab2:
    st.dataframe(res)
    st.dataframe(foco[cols])
    st.download_button("⬇️ Exportar a Excel", a_excel(res, foco[cols], descartados),
                       file_name=f"SIC_{hoy:%Y%m%d}.xlsx")

with tab3:
    st.caption("Filas leídas de la hoja que no se cuentan (sin fecha válida o sin Status).")
    st.dataframe(descartados)
