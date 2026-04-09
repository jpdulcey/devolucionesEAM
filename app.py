import io
import re
import base64
from datetime import datetime

import gspread
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from docx import Document
from docx.shared import Inches
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(page_title="Calidad analistas", layout="wide")

# HEADER PRO
col1, col2 = st.columns([5, 2])

with col1:
    st.markdown(
        "<h1 style='margin-bottom: 0;'>DEVOLUCIONES</h1><hr>",
        unsafe_allow_html=True
    )

with col2:
    st.image("logo_DANE.jpg", use_container_width=True)

PUNTAJE_BASE = 100
SHEET_ID = "1xrDybkfOPlH3fLHEedPQG77Sf3PfUQzC7wKXPTber_g"


def limpiar_nombre_archivo(texto):
    texto = str(texto).strip()
    texto = re.sub(r'[\\/*?:"<>|]', "", texto)
    texto = re.sub(r"\s+", "_", texto)
    return texto


def conectar_sheet():
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]

    creds = ServiceAccountCredentials.from_json_keyfile_dict(
        dict(st.secrets["gcp_service_account"]), scope
    )

    client = gspread.authorize(creds)
    return client.open_by_key(SHEET_ID).sheet1


def guardar_registro_sheet(*row):
    try:
        conectar_sheet().append_row(list(row), value_input_option="USER_ENTERED")
        return True, None
    except Exception as e:
        return False, str(e)


@st.cache_data
def cargar_fuentes():
    df_control = pd.read_excel("control.xlsx")
    df_capdirest = pd.read_excel("capdirest.xlsx")

    df_control.columns = df_control.columns.str.lower()
    df_capdirest.columns = df_capdirest.columns.str.lower()

    return df_control.merge(df_capdirest, on="nordest", how="left")


@st.cache_data
def cargar_puntajes():
    df = pd.read_excel("puntajes.xlsx")
    df["orden"] = range(len(df))
    return df


def pegar_imagen_component(key):
    html = f"""
    <textarea id="paste-{key}" style="width:100%;height:70px;"></textarea>
    <div id="preview-{key}"></div>

    <script>
    const area = document.getElementById("paste-{key}");
    const preview = document.getElementById("preview-{key}");

    area.addEventListener("paste", e => {{
        const items = e.clipboardData.items;
        for (let i=0;i<items.length;i++) {{
            if (items[i].type.includes("image")) {{
                const file = items[i].getAsFile();
                const reader = new FileReader();

                reader.onload = ev => {{
                    preview.innerHTML = `<img src="${{ev.target.result}}" width="200"/>`;
                    window.parent.postMessage({{
                        type:"streamlit:setComponentValue",
                        value: ev.target.result
                    }},"*");
                }};
                reader.readAsDataURL(file);
            }}
        }}
    }});
    </script>
    """
    return components.html(html, height=150)


def generar_word(data, info):
    doc = Document()

    doc.add_heading("Resultado de Evaluación", 1)

    for k, v in info.items():
        doc.add_paragraph(f"{k}: {v}")

    doc.add_heading("Observaciones", 2)

    for item in data:
        doc.add_paragraph(item["subtitulo"]).bold = True
        doc.add_paragraph(item["texto"])

        if item["imagen"]:
            doc.add_picture(io.BytesIO(item["imagen"]), width=Inches(4.5))

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


df_base = cargar_fuentes()
df_puntajes = cargar_puntajes()

st.subheader("Identificación")
nordest = st.text_input("NORDEST")

if nordest:
    fila = df_base[df_base["nordest"] == nordest]

    if not fila.empty:
        f = fila.iloc[0]

        st.write("Analista:", f.get("usuario", ""))
        st.write("Establecimiento:", f.get("nomest", ""))

        seleccionados = []

        st.subheader("Evaluación")

        for _, r in df_puntajes.iterrows():
            idx = r["orden"]

            check = st.checkbox(f"{r['SUBTÍTULO_2']} ({r['PUNTAJE']})", key=f"c{idx}")

            if check:
                texto = st.text_area("Observación", key=f"t{idx}")

                img = pegar_imagen_component(idx)

                imagen_bytes = None
                if img:
                    try:
                        _, enc = img.split(",", 1)
                        imagen_bytes = base64.b64decode(enc)
                    except:
                        pass

                seleccionados.append({
                    "subtitulo": r["SUBTÍTULO_2"],
                    "puntaje": r["PUNTAJE"],
                    "texto": texto,
                    "imagen": imagen_bytes
                })

        if st.button("Finalizar"):
            puntaje = max(0, 100 - sum(x["puntaje"] for x in seleccionados))

            rec = "DEVOLVER" if puntaje < 90 else "ENVIAR CORREO"

            st.write("Puntaje:", puntaje)
            st.write("Recomendación:", rec)

            decision = st.radio("Decisión final", ["ENVIAR CORREO", "DEVOLVER"])

            if st.button("Confirmar decisión"):
                ok, err = guardar_registro_sheet(
                    datetime.now(),
                    f.get("usuario", ""),
                    f.get("nordemp", ""),
                    nordest,
                    puntaje,
                    decision
                )

                if ok:
                    st.success("Guardado ✔")
                else:
                    st.error(err)

                doc = generar_word(seleccionados, {
                    "NORDEST": nordest,
                    "Resultado": decision,
                    "Puntaje": puntaje
                })

                st.download_button("Descargar Word", doc, "resultado.docx")

    else:
        st.warning("No encontrado")
