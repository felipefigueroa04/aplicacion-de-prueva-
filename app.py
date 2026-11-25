from io import BytesIO
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd
import plotly.express as px
import streamlit as st


NOMBRE_HOJA = "Recursos sin procesar"
COLUMNAS_REQUERIDAS = [
    "WBS",
    "Grupo1",
    "Grupo2",
    "Grupo3",
    "Elemento",
    "Objeto",
    "Recurso",
    "CantidadPrimaria",
]

st.set_page_config(page_title="Análisis de recursos de obra", layout="wide")


def clasificar_recurso(recurso):
    """Clasifica el tipo de recurso a partir de su prefijo."""
    prefijo = str(recurso).strip().upper()
    if prefijo.startswith("MO-"):
        return "MO"
    if prefijo.startswith("MAQ-"):
        return "MAQ"
    if prefijo.startswith("MAT-"):
        return "MAT"
    return "Otros"


def crear_excel(resumen, detalle):
    """Genera un libro Excel en memoria con las dos tablas descargables."""
    archivo = BytesIO()
    with pd.ExcelWriter(archivo, engine="openpyxl") as writer:
        resumen.to_excel(writer, sheet_name="Resumen_Duracion", index=False)
        detalle.to_excel(writer, sheet_name="Detalle_Recursos", index=False)
        for hoja in writer.sheets.values():
            hoja.freeze_panes = "A2"
            hoja.auto_filter.ref = hoja.dimensions
            for fila in hoja.iter_rows(min_row=2):
                for celda in fila:
                    if isinstance(celda.value, (int, float)):
                        celda.number_format = "#,##0.00"
    archivo.seek(0)
    return archivo.getvalue()


def formato_numero(valor):
    return f"{valor:,.2f}"


st.title("Análisis de tiempos y recursos de obra")
st.caption(
    "Explora los días-recurso, estima duraciones por partida y exporta los resultados."
)

with st.sidebar:
    st.header("Datos y filtros")
    archivo_subido = st.file_uploader(
        "Subir archivo Excel (.xlsx)", type=["xlsx"]
    )
    archivo_muestra = Path(__file__).parent / "Avance 1 Prisma.xlsx"

if archivo_subido is not None:
    origen_excel = archivo_subido
    nombre_archivo = archivo_subido.name
elif archivo_muestra.exists():
    origen_excel = archivo_muestra
    nombre_archivo = archivo_muestra.name
    st.sidebar.info(
        f"Usando el archivo local **{nombre_archivo}**. "
        "Puedes subir otro Excel para reemplazarlo."
    )
else:
    st.info("Sube un archivo .xlsx para comenzar el análisis.")
    st.stop()

try:
    libro = pd.ExcelFile(origen_excel, engine="openpyxl")
    if NOMBRE_HOJA not in libro.sheet_names:
        st.error(
            f"El archivo **{nombre_archivo}** no contiene la hoja "
            f"'{NOMBRE_HOJA}'. Hojas disponibles: {', '.join(libro.sheet_names)}."
        )
        st.stop()
    datos = pd.read_excel(
        libro, sheet_name=NOMBRE_HOJA, engine="openpyxl"
    )
except (ValueError, OSError) as error:
    st.error(f"No se pudo leer el archivo Excel: {error}")
    st.stop()

faltantes = [columna for columna in COLUMNAS_REQUERIDAS if columna not in datos.columns]
if faltantes:
    st.error(
        "Faltan columnas requeridas en la hoja "
        f"'{NOMBRE_HOJA}': {', '.join(faltantes)}."
    )
    st.stop()
if datos.empty:
    st.warning(f"La hoja '{NOMBRE_HOJA}' no contiene filas para analizar.")
    st.stop()

# Normaliza texto y valores numéricos antes de aplicar la lógica de recursos.
for columna in ["WBS", "Grupo1", "Grupo2", "Grupo3", "Elemento", "Objeto", "Recurso"]:
    datos[columna] = datos[columna].fillna("").astype(str).str.strip()
datos["Grupo3"] = datos["Grupo3"].replace("", "General")
for columna in ["WBS", "Grupo1", "Grupo2", "Elemento", "Objeto", "Recurso"]:
    datos[columna] = datos[columna].replace("", "Sin especificar")
datos["CantidadPrimaria"] = pd.to_numeric(
    datos["CantidadPrimaria"], errors="coerce"
).fillna(0)
datos["Tipo_Recurso"] = datos["Recurso"].apply(clasificar_recurso)

# Se consideran recursos de tiempo los marcados en días y todos los MO/MAQ.
es_tiempo = datos["Recurso"].str.contains(r"\(d[ií]as\)", case=False, na=False)
datos_tiempo = datos.loc[es_tiempo | datos["Tipo_Recurso"].isin(["MO", "MAQ"])].copy()

if datos_tiempo.empty:
    st.warning("No se encontraron recursos de tiempo (MO, MAQ o marcados en días).")
    st.stop()

with st.sidebar:
    cuadrillas = st.slider(
        "Cuadrillas / Frentes simultáneos",
        min_value=0.1,
        max_value=20.0,
        value=1.0,
        step=0.1,
    )
    modo_duracion = st.radio(
        "Cálculo de duración",
        ["Duración Crítica", "Suma Total de Días"],
        help=(
            "La duración crítica toma el máximo de días entre los recursos de "
            "cada partida; la suma total agrega todos sus días-recurso."
        ),
    )

    grupos_disponibles = sorted(datos_tiempo["Grupo1"].unique().tolist())
    grupos_seleccionados = st.multiselect(
        "Grupo1 (Capítulo)",
        grupos_disponibles,
        default=grupos_disponibles,
    )
    base_filtro = datos_tiempo[
        datos_tiempo["Grupo1"].isin(grupos_seleccionados)
    ]
    partidas_disponibles = sorted(base_filtro["Grupo2"].unique().tolist())
    partidas_seleccionadas = st.multiselect(
        "Grupo2 (Partida)",
        partidas_disponibles,
        default=partidas_disponibles,
    )
    base_filtro = base_filtro[
        base_filtro["Grupo2"].isin(partidas_seleccionadas)
    ]
    tipos_disponibles = sorted(base_filtro["Tipo_Recurso"].unique().tolist())
    tipos_seleccionados = st.multiselect(
        "Tipo_Recurso",
        tipos_disponibles,
        default=tipos_disponibles,
    )

filtrado = base_filtro[base_filtro["Tipo_Recurso"].isin(tipos_seleccionados)].copy()
if filtrado.empty:
    st.warning("No hay datos para la combinación de filtros seleccionada.")
    st.stop()

claves_partida = ["Grupo1", "Grupo2"]
claves_recurso = claves_partida + ["Recurso", "Tipo_Recurso"]
recursos_agrupados = (
    filtrado.groupby(claves_recurso, as_index=False, dropna=False)["CantidadPrimaria"]
    .sum()
    .rename(columns={"CantidadPrimaria": "Dias_Recurso"})
)

def combinar_valores(serie):
    valores = sorted({str(valor) for valor in serie if str(valor).strip()})
    return ", ".join(valores)


resumen_partida = (
    filtrado.groupby(claves_partida, as_index=False, dropna=False)
    .agg(
        WBS=("WBS", combinar_valores),
        Elementos=("Elemento", combinar_valores),
    )
)
metricas_duracion = (
    recursos_agrupados.groupby(claves_partida, as_index=False)
    .agg(
        Dias_Criticos=("Dias_Recurso", "max"),
        Dias_Totales=("Dias_Recurso", "sum"),
    )
)
resumen_partida = resumen_partida.merge(
    metricas_duracion, on=claves_partida, how="left"
)
columna_duracion = (
    "Dias_Criticos" if modo_duracion == "Duración Crítica" else "Dias_Totales"
)
resumen_partida["Duracion_Estimada"] = (
    resumen_partida[columna_duracion] / cuadrillas
)
resumen_partida["Partida"] = resumen_partida["Grupo2"]
resumen_partida = resumen_partida.sort_values(
    "Duracion_Estimada", ascending=False
).reset_index(drop=True)

total_dias = resumen_partida["Duracion_Estimada"].sum()
partida_critica = resumen_partida.iloc[0]
recursos_unicos = filtrado["Recurso"].nunique()

col_kpi1, col_kpi2, col_kpi3, col_kpi4 = st.columns(4)
col_kpi1.metric("Días totales estimados", formato_numero(total_dias))
col_kpi2.metric(
    "Partida más crítica",
    str(partida_critica["Grupo2"]),
    f"{formato_numero(partida_critica['Duracion_Estimada'])} días",
)
col_kpi3.metric("Total de partidas", f"{len(resumen_partida):,}")
col_kpi4.metric("Recursos únicos", f"{recursos_unicos:,}")
st.caption(
    f"Archivo: {nombre_archivo} · Cuadrillas: {cuadrillas:.1f} · "
    f"Método: {modo_duracion}"
)

st.subheader("Duración estimada por partida")
fig_duracion = px.bar(
    resumen_partida,
    x="Duracion_Estimada",
    y="Partida",
    color="Grupo1",
    orientation="h",
    text="Duracion_Estimada",
    hover_data={
        "Grupo1": True,
        "WBS": True,
        "Duracion_Estimada": ":.2f",
    },
    labels={
        "Duracion_Estimada": "Duración (días)",
        "Partida": "Grupo2 (Partida)",
        "Grupo1": "Capítulo",
    },
)
fig_duracion.update_traces(texttemplate="%{text:.2f}", textposition="outside")
fig_duracion.update_layout(
    yaxis={"categoryorder": "total ascending"},
    xaxis_title="Duración (días)",
    legend_title="Grupo1",
    margin={"l": 10, "r": 30, "t": 20, "b": 10},
)
st.plotly_chart(fig_duracion, use_container_width=True)

col_arbol, col_comparativa = st.columns(2)
with col_arbol:
    st.subheader("Distribución jerárquica de recursos")
    datos_arbol = filtrado.loc[filtrado["CantidadPrimaria"] > 0]
    if datos_arbol.empty:
        st.info("No hay cantidades positivas para construir el gráfico jerárquico.")
    else:
        fig_arbol = px.treemap(
            datos_arbol,
            path=["Grupo1", "Grupo2", "Elemento", "Recurso"],
            values="CantidadPrimaria",
            color="Tipo_Recurso",
            hover_data={"Objeto": True, "CantidadPrimaria": ":.2f"},
            labels={
                "CantidadPrimaria": "Días / cantidad",
                "Tipo_Recurso": "Tipo de recurso",
            },
        )
        fig_arbol.update_layout(margin={"l": 0, "r": 0, "t": 10, "b": 0})
        st.plotly_chart(fig_arbol, use_container_width=True)

with col_comparativa:
    st.subheader("Días de Mano de Obra vs Maquinaria")
    comparativa = recursos_agrupados[
        recursos_agrupados["Tipo_Recurso"].isin(["MO", "MAQ"])
    ].copy()
    comparativa["Dias_Ajustados"] = comparativa["Dias_Recurso"] / cuadrillas
    comparativa = (
        comparativa.groupby(claves_partida + ["Tipo_Recurso"], as_index=False)[
            "Dias_Ajustados"
        ]
        .sum()
    )
    pivote = comparativa.pivot_table(
        index=claves_partida,
        columns="Tipo_Recurso",
        values="Dias_Ajustados",
        aggfunc="sum",
        fill_value=0,
    )
    for tipo in ["MO", "MAQ"]:
        if tipo not in pivote.columns:
            pivote[tipo] = 0.0
    pivote["Etiqueta"] = [f"{grupo2} · {grupo1}" for grupo1, grupo2 in pivote.index]
    pivote = pivote.sort_values("MO", ascending=True)

    alto_figura = max(4, min(28, 0.35 * len(pivote)))
    figura, eje = plt.subplots(figsize=(10, alto_figura))
    posiciones = range(len(pivote))
    eje.barh(
        [posicion - 0.18 for posicion in posiciones],
        pivote["MO"],
        height=0.35,
        label="MO",
        color="#2a9d8f",
    )
    eje.barh(
        [posicion + 0.18 for posicion in posiciones],
        pivote["MAQ"],
        height=0.35,
        label="MAQ",
        color="#e76f51",
    )
    eje.set_yticks(list(posiciones))
    eje.set_yticklabels(pivote["Etiqueta"], fontsize=8)
    eje.set_xlabel("Días estimados")
    eje.set_ylabel("Partida")
    eje.legend()
    eje.grid(axis="x", linestyle=":", alpha=0.4)
    figura.tight_layout()
    st.pyplot(figura, use_container_width=True)

    imagen_png = BytesIO()
    figura.savefig(imagen_png, format="png", dpi=300, bbox_inches="tight")
    plt.close(figura)
    imagen_png.seek(0)
    st.download_button(
        "Descargar gráfico comparativo (PNG, 300 DPI)",
        data=imagen_png.getvalue(),
        file_name="dias_mo_vs_maq.png",
        mime="image/png",
    )

st.subheader("Desglose por partida")
for _, fila in resumen_partida.iterrows():
    etiqueta = f"{fila['Grupo2']} · {fila['Grupo1']}"
    with st.expander(
        f"{etiqueta} — {formato_numero(fila['Duracion_Estimada'])} días"
    ):
        st.write(f"**WBS:** {fila['WBS']}")
        st.write(
            f"**Duración estimada:** "
            f"{formato_numero(fila['Duracion_Estimada'])} días"
        )
        detalle_partida = (
            filtrado[
                (filtrado["Grupo1"] == fila["Grupo1"])
                & (filtrado["Grupo2"] == fila["Grupo2"])
            ]
            .groupby(["Elemento", "Recurso", "Tipo_Recurso"], as_index=False)[
                "CantidadPrimaria"
            ]
            .sum()
            .rename(columns={"CantidadPrimaria": "Cantidad / días"})
        )
        st.dataframe(
            detalle_partida.style.format({"Cantidad / días": "{:.2f}"}),
            use_container_width=True,
            hide_index=True,
        )

st.subheader("Exportar resultados")
detalle_exportacion = filtrado[
    [
        "WBS",
        "Grupo1",
        "Grupo2",
        "Grupo3",
        "Elemento",
        "Objeto",
        "Recurso",
        "Tipo_Recurso",
        "CantidadPrimaria",
    ]
].copy()
detalle_exportacion = (
    detalle_exportacion.groupby(
        [
            "WBS",
            "Grupo1",
            "Grupo2",
            "Grupo3",
            "Elemento",
            "Objeto",
            "Recurso",
            "Tipo_Recurso",
        ],
        as_index=False,
        dropna=False,
    )["CantidadPrimaria"]
    .sum()
)
detalle_exportacion["Dias_Recurso"] = detalle_exportacion["CantidadPrimaria"]
resumen_exportacion = resumen_partida[
    [
        "WBS",
        "Grupo1",
        "Grupo2",
        "Dias_Criticos",
        "Dias_Totales",
        "Duracion_Estimada",
    ]
].copy()
resumen_exportacion["Metodo_Duracion"] = modo_duracion
resumen_exportacion["Cuadrillas"] = cuadrillas
datos_excel = crear_excel(resumen_exportacion, detalle_exportacion)
st.download_button(
    "Descargar análisis en Excel",
    data=datos_excel,
    file_name="analisis_recursos_obra.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
