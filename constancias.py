import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime
import tempfile
import os
#from docx2pdf import convert

st.set_page_config(page_title="Constancias UAEM", page_icon="📄")

with st.sidebar:
    st.image("uaem logo.png", width=200)
    st.markdown("### UAEM Valle de Chalco")
    st.write("Herramienta Administrativa")
    st.markdown("---")

st.markdown("""
<style>
[data-testid="stAppViewContainer"] {
    background: linear-gradient(135deg, #e8f5e9, #fdf6e3);
}
[data-testid="stSidebar"] {
    background-color: #fdf6e3;
}
h1 {
    color: #004a2f !important;
    border-bottom: 3px solid #b3a369;
}
h3 {
    color: #b3a369 !important;
}
div.stButton > button {
    background-color: #004a2f !important;
    color: white !important;
    border: 2px solid #b3a369 !important;
    border-radius: 10px;
}
div.stButton > button:hover {
    background-color: #b3a369 !important;
    color: #004a2f !important;
}
input, textarea {
    background-color: #ffffff !important;
    color: #000000 !important;
    border: 2px solid #b3a369 !important;
    border-radius: 8px !important;
    padding: 6px !important;
}
[data-testid="stStatusWidget"] {
    display: none;
}
</style>
""", unsafe_allow_html=True)

def limpiar_campos():
    st.session_state["nombre"] = ""
    st.session_state["motivo"] = ""
    st.session_state["motivo_guardado"] = ""
    if "fecha_manual" in st.session_state:
        st.session_state["fecha_manual"] = ""

def reemplazar_nombre(doc, nombre):
    for p in doc.paragraphs:
        if "{{NOMBRE}}" in p.text:
            p.text = ""
            run = p.add_run(nombre)
            run.bold = True
            run.font.size = Pt(16)
            run.font.name = "Arial"
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

def reemplazar_motivo(doc, motivo):
    for p in doc.paragraphs:
        if "{{MOTIVO}}" in p.text:
            p.text = ""
            run = p.add_run(motivo)
            run.font.size = Pt(12)
            run.font.name = "Arial"
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                if "{{MOTIVO}}" in celda.text:
                    celda.text = motivo

def reemplazar_fecha(doc, fecha):
    for p in doc.paragraphs:
        if "{{FECHA}}" in p.text:
            p.text = p.text.replace("{{FECHA}}", fecha)
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                if "{{FECHA}}" in celda.text:
                    celda.text = celda.text.replace("{{FECHA}}", fecha)

if "nombre" not in st.session_state:
    st.session_state["nombre"] = ""
if "motivo" not in st.session_state:
    st.session_state["motivo"] = ""
if "motivo_guardado" not in st.session_state:
    st.session_state["motivo_guardado"] = ""

st.title("Generador de Constancias")
st.subheader("Área de Investigación y Estudios Avanzados")

tipo_constancia = st.radio("Tipo de constancia", ["CON FIRMA DE DIRECTORA", "SIN FIRMA DE DIRECTORA"])

nombre = st.text_input("Nombre:", key="nombre")

bloquear = st.checkbox("Bloquear motivo")

if not bloquear:
    motivo = st.text_area("Motivo:", key="motivo")
    if motivo:
        st.session_state["motivo_guardado"] = motivo
else:
    motivo = st.session_state["motivo_guardado"]
    st.text_area("Motivo:", value=motivo, disabled=True)

modo_fecha = st.radio("Tipo de fecha", ["Automatica", "Manual"])

if modo_fecha == "Manual":
    fecha = st.text_input("Fecha:", key="fecha_manual")
else:
    meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio",
             "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
    hoy = datetime.now()
    fecha = f"a {hoy.day} de {meses[hoy.month-1]} de {hoy.year}"

st.button("Limpiar todos los campos", on_click=limpiar_campos)

if st.button("Generar constancia"):
    if not nombre.strip() or not motivo.strip() or not fecha.strip():
        st.warning("Completa todos los campos")
    else:
        with st.spinner("Generando constancia..."):
            try:
                BASE_DIR = os.path.dirname(__file__)
                if tipo_constancia == "CON FIRMA DE DIRECTORA":
                    ruta = os.path.join(BASE_DIR, "CON FIRMA DE DIRECTORA.docx")
                else:
                    ruta = os.path.join(BASE_DIR, "SIN FIRMA DE DIRECTORA.docx")

                doc = Document(ruta)

                reemplazar_nombre(doc, nombre)
                reemplazar_motivo(doc, motivo)
                reemplazar_fecha(doc, fecha)

                with tempfile.TemporaryDirectory() as tmpdir:
                    docx_path = os.path.join(tmpdir, "temp.docx")
                    pdf_path = os.path.join(tmpdir, "temp.pdf")

                    doc.save(docx_path)

                    try:
                        convert(docx_path, pdf_path)
                        with open(pdf_path, "rb") as f:
                            pdf_bytes = f.read()

                        st.success("Constancia generada en PDF")

                        st.download_button(
                            label="Descargar PDF",
                            data=pdf_bytes,
                            file_name=f"Constancia_{nombre}.pdf",
                            mime="application/pdf"
                        )
                    except:
                        with open(docx_path, "rb") as f:
                            docx_bytes = f.read()

                        st.warning("No se pudo generar PDF en la nube")

                        st.download_button(
                            label="Descargar Word",
                            data=docx_bytes,
                            file_name=f"Constancia_{nombre}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
            except Exception as e:
                st.error(f"Error: {e}")
