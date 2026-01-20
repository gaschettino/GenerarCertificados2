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

# --- Configuración de la página ---
st.set_page_config(page_title="Generador de Certificados", layout="centered")

st.title("Generador de Certificados")
st.write("Cualquier consulta enviar mail a gaschettino@garrahan.gov.ar")

# --- Subida de archivos ---
uploaded_template = st.file_uploader(
    "Subí el template del certificado (.pptx)", type=["pptx"]
)
uploaded_excel = st.file_uploader(
    "Subí el listado de asistentes (.xlsx). "
    "Tiene que tener dos columnas: 'Nombre' y 'Apellido'"
)

# --- Inicializar session_state ---
if "color_mode" not in st.session_state:
    st.session_state.color_mode = "predefinido"

if "color_predefinido" not in st.session_state:
    st.session_state.color_predefinido = "Negro"

if "rgb_input" not in st.session_state:
    st.session_state.rgb_input = ""

if "hex_input" not in st.session_state:
    st.session_state.hex_input = ""

# --- Opciones de formato ---
st.subheader("Formato del nombre")

st.info(
    "Las fuentes disponibles dependen del servidor donde corre la aplicación. "
    "Si una fuente no está instalada, se usará una alternativa automáticamente."
)

# --- Fuentes seguras ---
fuentes_disponibles = [
    "DejaVu Sans",
    "DejaVu Serif",
    "Liberation Sans",
    "Liberation Serif"
]

fuente_seleccionada = st.selectbox(
    "Tipo de fuente",
    fuentes_disponibles,
    index=0
)

# --- Colores ---
st.subheader("Color de la letra")

st.markdown(
    "Podés elegir un color predefinido o ingresar un color personalizado (RGB o HEX).  \n"
    "👉 [Ver códigos de color](https://htmlcolorcodes.com/)"
)

colores_disponibles = {
    "Negro": RGBColor(0, 0, 0),
    "Azul": RGBColor(0, 0, 180),
    "Rojo": RGBColor(180, 0, 0),
    "Verde": RGBColor(0, 140, 0),
    "Gris": RGBColor(90, 90, 90)
}

# --- Selector de modo ---
color_mode = st.radio(
    "Modo de selección de color",
    ["predefinido", "rgb", "hex"],
    index=["predefinido", "rgb", "hex"].index(st.session_state.color_mode),
    horizontal=True
)

st.session_state.color_mode = color_mode

rgb_personalizado = None
hex_personalizado = None

# --- Predefinido ---
if color_mode == "predefinido":
    color_predefinido = st.selectbox(
        "Color predefinido",
        list(colores_disponibles.keys()),
        index=list(colores_disponibles.keys()).index(st.session_state.color_predefinido)
    )
    st.session_state.color_predefinido = color_predefinido

# --- RGB ---
if color_mode == "rgb":
    rgb_input = st.text_input(
        "Ingresá RGB (ej: 34,139,34)",
        value=st.session_state.rgb_input,
        placeholder="R,G,B"
    )
    st.session_state.rgb_input = rgb_input

    try:
        r, g, b = [int(x.strip()) for x in rgb_input.split(",")]
        if all(0 <= v <= 255 for v in (r, g, b)):
            rgb_personalizado = RGBColor(r, g, b)
        else:
            st.warning("Los valores RGB deben estar entre 0 y 255.")
    except:
        if rgb_input:
            st.warning("Formato inválido. Usá: R,G,B (ej: 255,0,0)")

# --- HEX ---
if color_mode == "hex":
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
        hex_personalizado = RGBColor(r, g, b)
    else:
        if hex_input:
            st.warning("Formato HEX inválido. Usá: #RRGGBB")

# --- Color final ---
if color_mode == "predefinido":
    color_seleccionado = colores_disponibles[st.session_state.color_predefinido]
elif color_mode == "rgb":
    color_seleccionado = rgb_personalizado or RGBColor(0, 0, 0)
elif color_mode == "hex":
    color_seleccionado = hex_personalizado or RGBColor(0, 0, 0)
else:
    color_seleccionado = RGBColor(0, 0, 0)

# --- Preview del color ---
st.markdown(
    f"<div style='width:120px;height:30px;border:1px solid #000;"
    f"background-color:rgb({color_seleccionado.rgb[0]},"
    f"{color_seleccionado.rgb[1]},"
    f"{color_seleccionado.rgb[2]});'></div>",
    unsafe_allow_html=True
)

# --- Función para convertir PPTX → PDF ---
def convert_to_pdf(input_pptx, output_dir):
    try:
        result = subprocess.run(
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

        if result.returncode == 0:
            os.remove(input_pptx)
            return True
        return False
    except Exception as e:
        print(f"Error al convertir {input_pptx}: {e}")
        return False

# --- Procesamiento ---
if uploaded_template and uploaded_excel:
    if st.button("🚀 Generar certificados"):
        with st.spinner("Generando certificados..."):
            with tempfile.TemporaryDirectory() as tmpdir:

                # Guardar template
                template_path = os.path.join(tmpdir, "template.pptx")
                with open(template_path, "wb") as f:
                    f.write(uploaded_template.read())

                # Leer Excel
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
                    st.error(
                        "No se encontró una columna válida "
                        "(Apellido/Nombre o Nombre y Apellido)."
                    )
                    st.stop()

                if "Asistió" in df.columns:
                    df = df[df["Asistió"].astype(str).str.upper() == "SI"]

                output_dir = os.path.join(tmpdir, "Certificados")
                os.makedirs(output_dir, exist_ok=True)

                nombres = df["Nombre y apellido"].dropna().unique()
                total = len(nombres)

                progress_bar = st.progress(0)
                status_text = st.empty()

                for i, nombre in enumerate(nombres):
                    status_text.text(
                        f"Procesando: {nombre} ({i+1}/{total})"
                    )

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
                                            run.font.color.rgb = color_seleccionado
                                    paragraph.alignment = PP_ALIGN.CENTER

                    safe_name = "".join(
                        c for c in nombre
                        if c.isalnum() or c in (" ", "-", "_")
                    ).rs
