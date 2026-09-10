from io import BytesIO

import pandas as pd
import streamlit as st


st.set_page_config(page_title="Hoja de verificación", page_icon="✅", layout="wide")

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


def tabla_inicial() -> pd.DataFrame:
    datos = {"Día": [f"Día {i}" for i in range(1, 13)], "Fecha": [None] * 12}
    for criterio in CRITERIOS:
        datos[criterio] = [False] * 12
    datos["Observaciones"] = [""] * 12
    return pd.DataFrame(datos)


def crear_excel(datos: pd.DataFrame) -> bytes:
    salida = BytesIO()
    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        datos.to_excel(writer, index=False, sheet_name="Verificación")
        hoja = writer.book["Verificación"]
        hoja.freeze_panes = "A2"
        for celda in hoja[1]:
            celda.font = __import__("openpyxl").styles.Font(bold=True)
        for columna in hoja.columns:
            letra = columna[0].column_letter
            ancho = max(len(str(c.value or "")) for c in columna) + 2
            hoja.column_dimensions[letra].width = min(ancho, 35)
    return salida.getvalue()


if "datos_verificacion" not in st.session_state:
    st.session_state.datos_verificacion = tabla_inicial()

st.title("✅ Hoja de verificación de inventario")
st.caption("Registra observaciones durante 12 días. Marca una casilla cuando la actividad haya sido observada o realizada.")

col_guardar, col_limpiar, col_exportar = st.columns([1, 1, 1.3])

with st.form("formulario_verificacion", border=False):
    configuracion = {
        "Día": st.column_config.TextColumn(disabled=True, width="small"),
        "Fecha": st.column_config.DateColumn(format="DD/MM/YYYY", width="medium"),
        "Observaciones": st.column_config.TextColumn(width="large"),
    }
    for criterio in CRITERIOS:
        configuracion[criterio] = st.column_config.CheckboxColumn(criterio, width="medium")

    editado = st.data_editor(
        st.session_state.datos_verificacion,
        column_config=configuracion,
        hide_index=True,
        use_container_width=True,
        num_rows="fixed",
        key="editor_verificacion",
    )
    guardar = st.form_submit_button("Guardar registros", type="primary", use_container_width=True)

if guardar:
    st.session_state.datos_verificacion = editado.copy()
    st.success("Registros guardados en esta sesión.")

with col_limpiar:
    if st.button("Limpiar registros", use_container_width=True):
        st.session_state.datos_verificacion = tabla_inicial()
        st.session_state.pop("editor_verificacion", None)
        st.rerun()

with col_exportar:
    archivo_excel = crear_excel(st.session_state.datos_verificacion)
    st.download_button(
        "Descargar Excel",
        data=archivo_excel,
        file_name="hoja_verificacion_inventario.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

st.info("Antes de descargar, pulsa “Guardar registros” para incluir las últimas modificaciones en el Excel.")
