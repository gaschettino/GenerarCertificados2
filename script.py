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
# Subida de archivos
# =========================
uploaded_template = st.file_uploader(
    "Subí el template del certificado (.pptx)", type=["pptx"]
)
uploaded_excel = st.file_uploader(
    "Subí el listado de asistentes (.xlsx). "
    "Tiene que tener dos columnas: 'Nombre' y 'Apellido'"
)

# =========================
# Estado inicial
# =========================
if "color_mode" not in st.session_state:
    st.session_state.color_mode = "predefinido"
if "color_predefinido" not in st.session_state:
    st.session_state.color_predefinido = "Negro"
if "rgb_input" not in st.session_state:
    st.session_state.rgb_input = ""
if "hex_input" not in st.session_state:
    st.session_state.hex_input = ""

# =========================
# Opciones de formato
# =========================
st.subheader("Formato del nombre")

st.info(
    "La vista previa es orientativa. "
    "El PDF final usa las fuentes disponibles en el servidor."
)

# --- Fuentes disponibles ---
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

# =========================
# Color de la letra
# =========================
st.subheader("Color de la letra")

st.markdown(
    "Elegí un color predefinido o ingresá uno personalizado (RGB o HEX).  \n"
    "👉 [Ver códigos de color](https://htmlcolorcodes.com/)"
)

# --- Colores predefinidos ---
colores_predefinidos = {
    "Negro": (0, 0, 0),
    "Azul": (0, 0, 180),
    "Rojo": (180, 0, 0),
    "Verde": (0, 140, 0),
    "Gris": (90, 90, 90)
}

color_mode = st.radio(
    "Modo de selección de color",
    ["predefinido", "rgb", "hex"],
    horizontal=True,
    index=["predefinido", "rgb", "hex"].index(st.session_state.color_mode)
)

st.session_state.color_mode = color_mode

# --- Color final (default NEGRO) ---
r, g, b = 0, 0, 0

# --- Predefinido ---
if color_mode == "predefinido":
    color_predefinido = st.selectbox(
        "Color predefinido",
        list(colores_predefinidos.keys()),
        index=list(colores_predefinidos.keys()).index(
            st.session_state.color_predefinido
        )
    )
    st.session_state.color_predefinido = color_predefinido
    r, g, b = colores_predefinidos[color_predefinido]

# --- RGB ---
elif color_mode == "rgb":
    rgb_input = st.text_input(
        "Ingresá RGB (ej: 34,139,34)",
        value=st.session_state.rgb_input,
        placeholder="R,G,B"
    )
    st.session_state.rgb_input = rgb_input

    try:
        r, g, b = [int(x.strip()) for x in rgb_input.split(",")]
        if not all(0 <= v <= 255 for v in (r, g, b)):
            st.warning("Los valores RGB deben estar entre 0 y 255.")
            r, g, b = 0, 0, 0
    except:
        if rgb_input:
            st.warning("Formato inválido. Usá: R,G,B (ej: 255,0,0)")
        r, g, b = 0, 0, 0

# --- HEX ---
elif color_mode == "hex":
    hex_input = st.text_input(
        "Ingresá HEX (ej: #228B22)",
        value=st.session_state.hex_input,
        placeholder="#RRGGBB"
    )
    st.session_state.hex_input = hex_input

    if re.match(r"^#([A-Fa-f0-9]{6})$", hex_input):
        r = int(hex_input[1:3], 16)
        g = int(hex_input[3:5], 16)
        b = int(hex_input[5:7], 16)
    else:
        if hex_input:
            st.warning("Formato HEX inválido. Usá: #RRGGBB")
        r, g, b = 0, 0, 0

# =========================
# Vista previa (YA con r,g,b definidos)
# =========================
st.subheader("Vista previa")

st.markdown(
    f"""
    <div style="
        font-family: '{fuente_seleccionada}', sans-serif;
        font-size: 28px;
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

# =========================
# Función PPTX → PDF
# =========================
def convert_to_pdf(input_pptx, output_dir):
    try:
        subprocess.run(
            [
                "libreoffice",
                "--headless",
                "--convert-to", "pdf",
                "--outdir", output_dir,
                input_pptx
            ],
            check=True,
            capture_output=True,
            text=True
        )
        os.remove(input_pptx)
        return True
    except Exception as e:
        print(f"Error al convertir {input_pptx}: {e}")
        return False

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

                if "Apellido" in df.columns and "Nombre" in df.columns:
                    df["Nombre y apellido"] = (
                        df["Apellido"].astype(str).str.strip() + " " +
                        df["Nombre"].astype(str).str.strip()
                    ).str.title()
                elif "Nombre Y Apellido" in df.columns:
                    df["Nombre y apellido"] = (
                        df["Nombre Y Apellido"].astype(str).str.title()
                    )
                else:
                    st.error("No se encontró columna de nombre válida.")
                    st.stop()

                if "Asistió" in df.columns:
                    df = df[df["Asistió"].astype(str).str.upper() == "SI"]

                output_dir = os.path.join(tmpdir, "Certificados")
                os.makedirs(output_dir, exist_ok=True)

                nombres = df["Nombre y apellido"].dropna().unique()
                total = len(nombres)

                progress = st.progress(0)
                status = st.empty()

                for i, nombre in enumerate(nombres):
                    status.text(f"Procesando: {nombre} ({i+1}/{total})")

                    prs = Presentation(template_path)

                    for slide in prs.slides:
                        for shape in slide.shapes:
                            if shape.has_text_frame:
                                for paragraph in shape.text_frame.paragraphs:
                                    for run in paragraph.runs:
                                        if "Nombre y apellido" in run.text:
                                            run.text = run.text.replace(
                                                "Nombre y apellido", nombre
                                            )
                                            run.font.name = fuente_seleccionada
                                            run.font.size = Pt(25)
                                            run.font.bold = True
                                            run.font.italic = True
                                            run.font.color.rgb = RGBColor(r, g, b)
                                    paragraph.alignment = PP_ALIGN.CENTER

                    safe_name = "".join(
                        c for c in nombre
                        if c.isalnum() or c in (" ", "-", "_")
                    ).rstrip().replace(" ", "_")

                    output_pptx = os.path.join(
                        output_dir, f"Certificado_{safe_name}.pptx"
                    )

                    prs.save(output_pptx)
                    convert_to_pdf(output_pptx, output_dir)

                    progress.progress((i + 1) / total)

                zip_path = os.path.join(tmpdir, "certificados.zip")
                shutil.make_archive(
                    zip_path.replace(".zip", ""),
                    "zip",
                    output_dir
                )

                with open(zip_path, "rb") as f:
                    st.download_button(
                        "Descargar Certificados PDF",
                        f,
                        "certificados.pdf.zip",
                        "application/zip"
                    )
