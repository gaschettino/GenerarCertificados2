import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
import os
import tempfile
import subprocess
import shutil
import re

# =========================
# Configuración de la página
# =========================
st.set_page_config(page_title="Generador de Certificados", layout="centered")

st.title("Generador de Certificados")
st.write("Cualquier consulta enviar mail a gaschettino@garrahan.gov.ar")

# =========================
# Subida de archivos (2 columnas)
# =========================
st.subheader("Archivos")

col1, col2 = st.columns(2)

with col1:
    uploaded_template = st.file_uploader(
        "Template del certificado (.pptx)",
        type=["pptx"]
    )

with col2:
    uploaded_excel = st.file_uploader(
        "Listado de asistentes (.xlsx)",
        type=["xlsx"]
    )

# =========================
# DNI
# =========================
st.subheader("Contenido del certificado")

incluye_dni = st.checkbox("El certificado incluye DNI")

if incluye_dni:
    st.info(
        "El Excel debe contener una columna llamada 'DNI' "
        "y el template debe incluir el texto 'Numero de DNI'."
    )

st.divider()

# =========================
# Formato del nombre y apellido
# =========================
st.subheader("Formato del nombre y apellido del paciente")

col_fuente, col_color = st.columns(2)

# ---- Columna izquierda: fuente + tamaño ----
with col_fuente:
    fuentes_disponibles = [
        "DejaVu Sans",
        "DejaVu Serif",
        "Liberation Sans",
        "Liberation Serif",
        "Arial",
        "Times New Roman",
        "Calibri"
    ]

    fuente_seleccionada = st.selectbox(
        "Tipo de fuente",
        fuentes_disponibles,
        index=0
    )

    tamaño_fuente = st.slider(
        "Tamaño de letra",
        min_value=12,
        max_value=50,
        value=25,
        step=1
    )

# ---- Columna derecha: color ----
with col_color:
    st.markdown(
        "Color de la letra  \n"
        "👉 [Ver códigos de color](https://htmlcolorcodes.com/)"
    )

    colores_predefinidos = {
        "Negro": (0, 0, 0),
        "Azul": (0, 0, 180),
        "Rojo": (180, 0, 0),
        "Verde": (0, 140, 0),
        "Gris": (90, 90, 90)
    }

    color_mode = st.radio(
        "Modo de selección",
        ["predefinido", "rgb", "hex"],
        horizontal=True
    )

    r, g, b = 0, 0, 0  # negro por default

    if color_mode == "predefinido":
        color_predefinido = st.selectbox(
            "Color predefinido",
            list(colores_predefinidos.keys())
        )
        r, g, b = colores_predefinidos[color_predefinido]

    elif color_mode == "rgb":
        rgb_input = st.text_input("RGB (ej: 34,139,34)")
        try:
            r, g, b = [int(x.strip()) for x in rgb_input.split(",")]
            if not all(0 <= v <= 255 for v in (r, g, b)):
                r, g, b = 0, 0, 0
        except:
            r, g, b = 0, 0, 0

    elif color_mode == "hex":
        hex_input = st.text_input("HEX (ej: #228B22)")
        if re.match(r"^#([A-Fa-f0-9]{6})$", hex_input):
            r = int(hex_input[1:3], 16)
            g = int(hex_input[3:5], 16)
            b = int(hex_input[5:7], 16)

st.divider()

# =========================
# Vista previa
# =========================
st.subheader("Vista previa")

st.markdown(
    f"""
    <div style="
        font-family: '{fuente_seleccionada}', sans-serif;
        font-size: {tamaño_fuente}px;
        font-weight: bold;
        font-style: italic;
        color: rgb({r},{g},{b});
        border: 1px dashed #999;
        padding: 12px;
        width: fit-content;
    ">
        Nombre y Apellido
    </div>
    """,
    unsafe_allow_html=True
)

st.caption(
    "La vista previa es orientativa. "
    "El PDF final usa las fuentes disponibles en el servidor."
)

st.divider()

# =========================
# Conversión PPTX → PDF
# =========================
def convert_to_pdf(input_pptx, output_dir):
    subprocess.run(
        [
            "libreoffice",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", output_dir,
            input_pptx
        ],
        check=True
    )
    os.remove(input_pptx)

# =========================
# Procesamiento
# =========================
if uploaded_template and uploaded_excel:
    if st.button("🚀 Generar certificados"):
        with st.spinner("Generando certificados..."):
            with tempfile.TemporaryDirectory() as tmpdir:

                template_path = os.path.join(tmpdir, "template.pptx")
                with open(template_path, "wb") as f:
                    f.write(uploaded_template.read())

                df = pd.read_excel(uploaded_excel)
                df.columns = df.columns.str.strip().str.title()

                if "Nombre" not in df.columns or "Apellido" not in df.columns:
                    st.error("El Excel debe tener columnas 'Nombre' y 'Apellido'.")
                    st.stop()

                if incluye_dni and "Dni" not in df.columns:
                    st.error("Marcaste que incluye DNI pero no existe la columna 'DNI'.")
                    st.stop()

                df["Nombre y apellido"] = (
                    df["Apellido"].astype(str).str.strip() + " " +
                    df["Nombre"].astype(str).str.strip()
                ).str.title()

                output_dir = os.path.join(tmpdir, "Certificados")
                os.makedirs(output_dir, exist_ok=True)

                for _, row in df.iterrows():
                    prs = Presentation(template_path)

                    for slide in prs.slides:
                        for shape in slide.shapes:
                            if shape.has_text_frame:
                                for paragraph in shape.text_frame.paragraphs:
                                    for run in paragraph.runs:

                                        if "Nombre y apellido" in run.text:
                                            run.text = run.text.replace(
                                                "Nombre y apellido",
                                                row["Nombre y apellido"]
                                            )
                                            run.font.name = fuente_seleccionada
                                            run.font.size = Pt(tamaño_fuente)
                                            run.font.bold = True
                                            run.font.italic = True
                                            run.font.color.rgb = RGBColor(r, g, b)

                                        if incluye_dni and "Numero de DNI" in run.text:
                                            run.text = run.text.replace(
                                                "Numero de DNI",
                                                str(row["Dni"])
                                            )

                                    paragraph.alignment = PP_ALIGN.CENTER

                    safe_name = row["Nombre y apellido"].replace(" ", "_")
                    pptx_path = os.path.join(
                        output_dir, f"Certificado_{safe_name}.pptx"
                    )

                    prs.save(pptx_path)
                    convert_to_pdf(pptx_path, output_dir)

                zip_path = os.path.join(tmpdir, "certificados.zip")
                shutil.make_archive(
                    zip_path.replace(".zip", ""),
                    "zip",
                    output_dir
                )

                with open(zip_path, "rb") as f:
                    st.download_button(
                        "Descargar certificados",
                        f,
                        "certificados.zip",
                        "application/zip"
                    )
