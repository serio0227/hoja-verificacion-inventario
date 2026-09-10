from io import BytesIO

import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill


# ---------------- CONFIGURACIÓN ----------------
st.set_page_config(
    page_title="Control de inventario",
    page_icon="📋",
    layout="wide",
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


# ---------------- FUNCIONES ----------------
def tabla_inicial():
    datos = {
        "Día": [f"Día {numero}" for numero in range(1, 13)],
        # Esta línea evita el error de Streamlit con la columna Fecha
        "Fecha": pd.Series([pd.NaT] * 12, dtype="datetime64[ns]"),
        "Observaciones": [""] * 12,
    }

    for criterio in CRITERIOS:
        datos[criterio] = [False] * 12

    return pd.DataFrame(datos)


def normalizar_fechas(datos):
    """Convierte la columna Fecha al formato compatible con Streamlit."""
    datos = datos.copy()
    datos["Fecha"] = pd.to_datetime(datos["Fecha"], errors="coerce")
    return datos


def crear_excel(datos):
    salida = BytesIO()

    with pd.ExcelWriter(salida, engine="openpyxl") as escritor:
        datos.to_excel(escritor, index=False, sheet_name="Verificación")

        hoja = escritor.book["Verificación"]
        hoja.freeze_panes = "A2"

        color_encabezado = PatternFill(
            start_color="123B5D",
            end_color="123B5D",
            fill_type="solid",
        )

        for celda in hoja[1]:
            celda.font = Font(bold=True, color="FFFFFF")
            celda.fill = color_encabezado
            celda.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True,
            )

        for fila in hoja.iter_rows(min_row=2):
            for celda in fila:
                celda.alignment = Alignment(
                    vertical="center",
                    wrap_text=True,
                )

        for columna in hoja.columns:
            letra = columna[0].column_letter
            ancho = max(len(str(celda.value or "")) for celda in columna) + 3
            hoja.column_dimensions[letra].width = min(ancho, 32)

        hoja.row_dimensions[1].height = 38

    return salida.getvalue()


# ---------------- ESTADO INICIAL ----------------
if "datos_verificacion" not in st.session_state:
    st.session_state.datos_verificacion = tabla_inicial()

# Convierte cualquier dato anterior de fecha para que no vuelva a generar error
st.session_state.datos_verificacion = normalizar_fechas(
    st.session_state.datos_verificacion
)


# ---------------- DISEÑO ----------------
st.markdown(
    """
    <style>
        .stApp {
            background: #f4f7fb;
        }

        .titulo-principal {
            color: #123b5d;
            font-size: 2.3rem;
            font-weight: 800;
            margin-bottom: 0;
        }

        .subtitulo {
            color: #587083;
            font-size: 1.08rem;
            margin-top: 0;
        }

        .tarjeta {
            background: white;
            border-radius: 16px;
            padding: 20px;
            border-left: 6px solid #1d7f70;
            box-shadow: 0 3px 10px rgba(18, 59, 93, 0.08);
            min-height: 115px;
        }

        .numero {
            font-size: 2rem;
            font-weight: 800;
            color: #123b5d;
            margin: 0;
        }

        .texto-tarjeta {
            color: #587083;
            font-size: 0.9rem;
            margin: 0;
        }

        div.stButton > button {
            border-radius: 9px;
            font-weight: 700;
            min-height: 44px;
        }

        div.stDownloadButton > button {
            border-radius: 9px;
            font-weight: 700;
            min-height: 44px;
            background-color: #147d6f;
            color: white;
            border: none;
        }

        @media (max-width: 700px) {
            .titulo-principal {
                font-size: 1.65rem;
            }

            .tarjeta {
                padding: 14px;
                min-height: 95px;
            }
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# ---------------- ENCABEZADO ----------------
st.markdown(
    '<p class="titulo-principal">📋 Hoja de verificación</p>',
    unsafe_allow_html=True,
)
st.markdown(
    '<p class="subtitulo">Control diario de inventario · Registro de 12 días</p>',
    unsafe_allow_html=True,
)

st.info(
    "Marca las casillas cuando la actividad haya sido observada o realizada. "
    "Puedes registrar observaciones y descargar el resultado en Excel."
)


# ---------------- INDICADORES ----------------
datos_actuales = st.session_state.datos_verificacion
total_casillas = len(datos_actuales) * len(CRITERIOS)
casillas_marcadas = int(datos_actuales[CRITERIOS].sum().sum())
avance = round((casillas_marcadas / total_casillas) * 100) if total_casillas else 0
dias_con_actividad = int(datos_actuales[CRITERIOS].any(axis=1).sum())

col1, col2, col3, col4 = st.columns(4)

with col1:
    st.markdown(
        f'<div class="tarjeta"><p class="numero">12</p>'
        f'<p class="texto-tarjeta">Días de verificación</p></div>',
        unsafe_allow_html=True,
    )

with col2:
    st.markdown(
        f'<div class="tarjeta"><p class="numero">{casillas_marcadas}</p>'
        f'<p class="texto-tarjeta">Actividades registradas</p></div>',
        unsafe_allow_html=True,
    )

with col3:
    st.markdown(
        f'<div class="tarjeta"><p class="numero">{dias_con_actividad}</p>'
        f'<p class="texto-tarjeta">Días con actividad</p></div>',
        unsafe_allow_html=True,
    )

with col4:
    st.markdown(
        f'<div class="tarjeta"><p class="numero">{avance}%</p>'
        f'<p class="texto-tarjeta">Avance de verificación</p></div>',
        unsafe_allow_html=True,
    )

st.write("")
st.subheader("Registro diario")


# ---------------- BOTÓN LIMPIAR ----------------
col_limpiar, col_espacio = st.columns([1, 4])

with col_limpiar:
    if st.button("🗑️ Limpiar registros", use_container_width=True):
        st.session_state.datos_verificacion = tabla_inicial()

        # Elimina el estado anterior del editor para mostrar la tabla limpia
        if "editor_verificacion" in st.session_state:
            del st.session_state["editor_verificacion"]

        st.rerun()


# ---------------- TABLA EDITABLE ----------------
datos_para_editar = normalizar_fechas(
    st.session_state.datos_verificacion
)

configuracion_columnas = {
    "Día": st.column_config.TextColumn(
        "Día",
        disabled=True,
        width="small",
    ),
    "Fecha": st.column_config.DateColumn(
        "Fecha",
        format="DD/MM/YYYY",
        width="medium",
    ),
    "Observaciones": st.column_config.TextColumn(
        "Observaciones",
        width="large",
    ),
}

for criterio in CRITERIOS:
    configuracion_columnas[criterio] = st.column_config.CheckboxColumn(
        criterio,
        help="Marca si esta actividad fue observada o realizada.",
        width="medium",
    )

datos_editados = st.data_editor(
    datos_para_editar,
    column_config=configuracion_columnas,
    hide_index=True,
    use_container_width=True,
    num_rows="fixed",
    key="editor_verificacion",
)

# Guarda los cambios realizados en la tabla
st.session_state.datos_verificacion = normalizar_fechas(datos_editados)


# ---------------- EXPORTACIÓN ----------------
st.write("")
st.subheader("Exportar información")

archivo_excel = crear_excel(st.session_state.datos_verificacion)

st.download_button(
    label="📥 Descargar reporte en Excel",
    data=archivo_excel,
    file_name="hoja_verificacion_inventario.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)

st.caption(
    "El archivo descargado contiene los 12 días, fechas, casillas marcadas y observaciones."
)
