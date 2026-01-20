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

# --- Opciones de formato ---
st.subheader("Formato del nombre")

st.info(
    "Las fuentes disponibles dependen del servidor donde corre la aplicación. "
    "Si una fuente no está instalada, se usará una alternativa automáticamente."
)

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

colores_disponibles = {
    "Negro": RGBColor(0, 0, 0),
    "Azul": RGBColor(0, 0, 180),
    "Rojo": RGBColor(180, 0, 0),
    "Verde": RGBColor(0, 140, 0),
    "Gris": RGBColor(90, 90, 90)
}

color_nombre = st.selectbox(
    "Color de la letra",
    list(colores_disponibles.keys())
)

color_seleccionado = colores_disponibles[color_nombre]

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
                    ).rstrip().replace(" ", "_")

                    output_pptx = os.path.join(
                        output_dir,
                        f"Certificado_{safe_name}.pptx"
                    )

                    prs.save(output_pptx)
                    convert_to_pdf(output_pptx, output_dir)

                    progress_bar.progress((i + 1) / total)

                status_text.text("Procesamiento completado")

                archivos = os.listdir(output_dir)
                pdf_count = len([f for f in archivos if f.endswith(".pdf")])

                st.success(f"Se generaron {pdf_count} certificados PDF.")

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
