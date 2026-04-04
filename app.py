import io
import re
from datetime import datetime

import pandas as pd
import streamlit as st
from docx import Document

st.set_page_config(page_title="Calidad analistas", layout="wide")
st.title("DEVOLUCIONES")

PUNTAJE_BASE = 100


def limpiar_nombre_archivo(texto):
    texto = str(texto).strip()
    texto = re.sub(r'[\\/*?:"<>|]', "", texto)
    texto = re.sub(r"\s+", "_", texto)
    return texto


@st.cache_data
def cargar_fuentes():
    df_control = pd.read_excel("control.xlsx")
    df_capdirest = pd.read_excel("capdirest.xlsx")

    df_control.columns = df_control.columns.str.strip().str.lower()
    df_capdirest.columns = df_capdirest.columns.str.strip().str.lower()

    df_control["nordest"] = df_control["nordest"].astype(str).str.strip()
    df_capdirest["nordest"] = df_capdirest["nordest"].astype(str).str.strip()

    if "usuarioss" not in df_control.columns:
        raise KeyError("No existe la columna 'usuarioss' en control.xlsx")
    if "nomest" not in df_capdirest.columns:
        raise KeyError("No existe la columna 'nomest' en capdirest.xlsx")
    if "nordemp" not in df_capdirest.columns:
        raise KeyError("No existe la columna 'nordemp' en capdirest.xlsx")

    df_control = df_control[["nordest", "usuarioss"]].drop_duplicates()
    df_capdirest = df_capdirest[["nordest", "nordemp", "nomest"]].drop_duplicates()

    df_base = df_control.merge(df_capdirest, on="nordest", how="outer")
    df_base = df_base.sort_values("nordest").reset_index(drop=True)

    return df_base


@st.cache_data
def cargar_puntajes():
    df = pd.read_excel("puntajes.xlsx")
    df.columns = [str(col).strip() for col in df.columns]

    columnas = ["Hoja", "TÍTULO", "SUBTÍTULO", "PUNTAJE"]
    for col in columnas:
        if col not in df.columns:
            raise KeyError(f"Falta la columna '{col}' en puntajes.xlsx")

    df = df[columnas].copy()
    df["orden"] = range(len(df))

    df["Hoja"] = pd.to_numeric(df["Hoja"], errors="coerce")
    df["TÍTULO"] = df["TÍTULO"].astype(str).str.strip()
    df["SUBTÍTULO"] = df["SUBTÍTULO"].astype(str).str.strip()
    df["PUNTAJE"] = pd.to_numeric(df["PUNTAJE"], errors="coerce").fillna(0)

    df = df.dropna(subset=["TÍTULO", "SUBTÍTULO"]).reset_index(drop=True)

    return df


def generar_word(nordemp, nordest, monitor, establecimiento, seleccionados, puntaje_final, decision):
    doc = Document()

    doc.add_heading("Resultado de Evaluación", 1)

    doc.add_paragraph(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    doc.add_paragraph(f"NORDEMP: {nordemp}")
    doc.add_paragraph(f"NORDEST: {nordest}")
    doc.add_paragraph(f"Monitor: {monitor}")
    doc.add_paragraph(f"Establecimiento: {establecimiento}")

    doc.add_heading("Resumen", 2)
    doc.add_paragraph(f"Puntaje final: {puntaje_final}")
    doc.add_paragraph(f"Resultado: {decision}")

    doc.add_heading("Observaciones", 2)

    if not seleccionados:
        doc.add_paragraph("No se registraron observaciones.")
    else:
        seleccionados_ordenados = sorted(seleccionados, key=lambda x: x["orden"])
        modulo_actual = None

        for item in seleccionados_ordenados:
            if item["titulo"] != modulo_actual:
                doc.add_heading(item["titulo"], 3)
                modulo_actual = item["titulo"]

            p_sub = doc.add_paragraph()
            run = p_sub.add_run(item["subtitulo"])
            run.bold = True

            if item["texto"]:
                doc.add_paragraph(item["texto"])
            else:
                doc.add_paragraph("")

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


df_base = cargar_fuentes()
df_puntajes = cargar_puntajes()

if "evaluacion_finalizada" not in st.session_state:
    st.session_state["evaluacion_finalizada"] = False

if "seleccionados_finales" not in st.session_state:
    st.session_state["seleccionados_finales"] = []

if "puntaje_final_final" not in st.session_state:
    st.session_state["puntaje_final_final"] = 100

if "decision_final" not in st.session_state:
    st.session_state["decision_final"] = ""

st.subheader("1. Identificación")

modo = st.radio("Modo de búsqueda", ["Escribir", "Seleccionar"], horizontal=True)

nordest = ""
if modo == "Escribir":
    nordest = st.text_input("NORDEST").strip()
else:
    lista = [""] + df_base["nordest"].dropna().astype(str).tolist()
    nordest = st.selectbox("NORDEST", lista)

if nordest:
    fila = df_base[df_base["nordest"] == nordest]

    if not fila.empty:
        monitor = fila.iloc[0]["usuarioss"]
        establecimiento = fila.iloc[0]["nomest"]
        nordemp = fila.iloc[0]["nordemp"]

        col1, col2, col3 = st.columns(3)
        with col1:
            st.text_input("NORDEMP", value=str(nordemp), disabled=True)
        with col2:
            st.text_input("Monitor", value=str(monitor), disabled=True)
        with col3:
            st.text_input("Establecimiento", value=str(establecimiento), disabled=True)

        st.divider()
        st.subheader("2. Evaluación")

        titulos_en_orden = df_puntajes["TÍTULO"].drop_duplicates().tolist()

        for titulo in titulos_en_orden:
            df_titulo = df_puntajes[df_puntajes["TÍTULO"] == titulo].copy()
            df_titulo = df_titulo.sort_values("orden")

            with st.expander(titulo, expanded=False):
                for _, fila2 in df_titulo.iterrows():
                    idx = int(fila2["orden"])

                    check_key = f"check_{idx}"
                    text_key = f"text_{idx}"

                    st.checkbox(
                        f"{fila2['SUBTÍTULO']} (-{fila2['PUNTAJE']})",
                        key=check_key
                    )

                    if st.session_state.get(check_key, False):
                        st.text_area(
                            f"Observación: {fila2['SUBTÍTULO']}",
                            key=text_key,
                            height=120,
                            placeholder="Aquí el analista puede escribir o pegar la observación..."
                        )

        st.divider()

        if st.button("Finalizar evaluación"):
            seleccionados = []

            for _, fila2 in df_puntajes.sort_values("orden").iterrows():
                idx = int(fila2["orden"])
                check_key = f"check_{idx}"
                text_key = f"text_{idx}"

                if st.session_state.get(check_key, False):
                    seleccionados.append({
                        "titulo": fila2["TÍTULO"],
                        "subtitulo": fila2["SUBTÍTULO"],
                        "puntaje": fila2["PUNTAJE"],
                        "texto": st.session_state.get(text_key, "").strip(),
                        "orden": idx
                    })

            puntaje_final = max(0, PUNTAJE_BASE - sum(x["puntaje"] for x in seleccionados))
            decision = "DEVOLVER" if puntaje_final < 90 else "ENVIAR CORREO"

            st.session_state["seleccionados_finales"] = seleccionados
            st.session_state["puntaje_final_final"] = puntaje_final
            st.session_state["decision_final"] = decision
            st.session_state["evaluacion_finalizada"] = True

        if st.session_state["evaluacion_finalizada"]:
            st.subheader("3. Resultado")

            c1, c2 = st.columns(2)
            c1.metric("Puntaje final", f"{st.session_state['puntaje_final_final']:g}")
            c2.metric("Resultado", st.session_state["decision_final"])

            if st.session_state["decision_final"] == "DEVOLVER":
                st.error("DEVOLVER")
            else:
                st.success("ENVIAR CORREO")

            archivo = generar_word(
                nordemp=nordemp,
                nordest=nordest,
                monitor=monitor,
                establecimiento=establecimiento,
                seleccionados=st.session_state["seleccionados_finales"],
                puntaje_final=st.session_state["puntaje_final_final"],
                decision=st.session_state["decision_final"]
            )

            nombre_archivo = (
                f"{limpiar_nombre_archivo(nordemp)}_"
                f"{limpiar_nombre_archivo(nordest)}_"
                f"{limpiar_nombre_archivo(establecimiento)}.docx"
            )

            st.download_button(
                "Descargar Word",
                archivo,
                file_name=nombre_archivo,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

    else:
        st.warning("No se encontró ese NORDEST.")