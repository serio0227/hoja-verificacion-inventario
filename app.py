from io import BytesIO

import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill


st.set_page_config(
    page_title="Control de Inventario",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="collapsed",
)

CRITERIOS = [
    "Entrada de material observada",
    "Entrada registrada",
    "Consumo/retiro observado",
    "Consumo/retiro registrado",
    "Salida de producto observada",
    "Salida registrada",
    "Consulta por memoria",
    "Conteo físico realizado",
]


def tabla_inicial():
    datos = {
        "Día": [f"Día {i}" for i in range(1, 13)],
        "Fecha": [None] * 12,
        "Observaciones": [""] * 12,
    }

    for criterio in CRITERIOS:
        datos[criterio] = [False] * 12

    return pd.DataFrame(datos)


def crear_excel(datos):
    salida = BytesIO()

    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        datos.to_excel(writer, index=False, sheet_name="Verificación")

        hoja = writer.book["Verificación"]
        hoja.freeze_panes = "A2"

        color_titulo = PatternFill("solid", fgColor="123B5D")

        for celda in hoja[1]:
            celda.font = Font(bold=True, color="FFFFFF")
            celda.fill = color_titulo
            celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for columna in hoja.columns:
            letra = columna[0].column_letter
            ancho = max(len(str(celda.value or "")) for celda in columna) + 3
            hoja.column_dimensions[letra].width = min(ancho, 34)

    return salida.getvalue()


def calcular_indicadores(datos):
    total_marcas = int(datos[CRITERIOS].sum().sum())
    entradas_observadas = int(datos["Entrada de material observada"].sum())
    salidas_observadas = int(datos["Salida de producto observada"].sum())
    conteos_realizados = int(datos["Conteo físico realizado"].sum())

    return total_marcas, entradas_observadas, salidas_observadas, conteos_realizados


if "datos_verificacion" not in st.session_state:
    st.session_state.datos_verificacion = tabla_inicial()


st.markdown(
    """
    <style>
        .stApp {
            background: #f4f7fb;
        }

        .bloque-principal {
            background: linear-gradient(135deg, #123b5d, #1f6f8b);
            padding: 30px;
            border-radius: 22px;
            color: white;
            margin-bottom: 22px;
            box-shadow: 0 12px 30px rgba(18, 59, 93, 0.20);
        }

        .bloque-principal h1 {
            margin: 0;
            font-size: 35px;
        }

        .bloque-principal p {
            margin: 8px 0 0;
            font-size: 16px;
            opacity: 0.92;
        }

        .tarjeta-info {
            background: white;
            border-radius: 16px;
            padding: 16px;
            border: 1px solid #dce5ec;
            box-shadow: 0 4px 12px rgba(18, 59, 93, 0.08);
        }

        div[data-testid="stMetric"] {
            background: white;
            border: 1px solid #dce5ec;
            border-radius: 16px;
            padding: 16px;
            box-shadow: 0 4px 12px rgba(18, 59, 93, 0.08);
        }

        div[data-testid="stDownloadButton"] button {
            background: #087f6b;
            color: white;
            border: none;
            font-weight: 700;
        }

        div[data-testid="stDownloadButton"] button:hover {
            background: #066b5a;
            color: white;
        }

        .stButton button {
            border-radius: 10px;
            font-weight: 700;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="bloque-principal">
        <h1>📦 Hoja de verificación de inventario</h1>
        <p>Registro diario de entradas, consumos, salidas y conteos físicos durante 12 días.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

total_marcas, entradas, salidas, conteos = calcular_indicadores(
    st.session_state.datos_verificacion
)

col1, col2, col3, col4 = st.columns(4)

col1.metric("✅ Marcas registradas", total_marcas)
col2.metric("📥 Entradas observadas", entradas)
col3.metric("📤 Salidas observadas", salidas)
col4.metric("📋 Conteos físicos", conteos)

st.markdown("### Registro de verificación")
st.caption("Marca las casillas cuando la actividad se haya observado o realizado.")

with st.form("formulario_verificacion"):
    configuracion = {
        "Día": st.column_config.TextColumn("Día", disabled=True, width="small"),
        "Fecha": st.column_config.DateColumn("Fecha", format="DD/MM/YYYY"),
        "Observaciones": st.column_config.TextColumn(
            "Observaciones",
            width="large",
            help="Escribe novedades o hallazgos importantes.",
        ),
    }

    for criterio in CRITERIOS:
        configuracion[criterio] = st.column_config.CheckboxColumn(
            criterio,
            width="medium",
        )

    datos_editados = st.data_editor(
        st.session_state.datos_verificacion,
        column_config=configuracion,
        hide_index=True,
        use_container_width=True,
        num_rows="fixed",
        key="editor_verificacion",
    )

    guardar = st.form_submit_button(
        "💾 Guardar registros",
        type="primary",
        use_container_width=True,
    )

if guardar:
    st.session_state.datos_verificacion = datos_editados.copy()
    st.success("Registros guardados correctamente.")

st.markdown("### Acciones")

accion1, accion2 = st.columns(2)

with accion1:
    if st.button("🗑️ Limpiar todos los registros", use_container_width=True):
        st.session_state.datos_verificacion = tabla_inicial()
        st.session_state.pop("editor_verificacion", None)
        st.rerun()

with accion2:
    archivo_excel = crear_excel(st.session_state.datos_verificacion)

    st.download_button(
        "⬇️ Descargar reporte en Excel",
        data=archivo_excel,
        file_name="hoja_verificacion_inventario.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

st.info(
    "Para incluir los últimos cambios en el Excel, primero pulsa “Guardar registros” "
    "y después descarga el reporte."
)
