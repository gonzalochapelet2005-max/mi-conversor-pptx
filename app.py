import streamlit as st
from pptx import Presentation
from pptx.util import Inches
import pytesseract
from PIL import Image
import io

st.set_page_config(page_title="Convertidor NotebookLM a PPTX", layout="centered")

st.title("🖼️ De Imagen a PowerPoint Editable")
st.write("Sube la imagen que te dio NotebookLM y te daré un archivo .pptx")

uploaded_file = st.file_uploader("Elige una imagen...", type=["jpg", "jpeg", "png"])

if uploaded_file is not None:
    image = Image.open(uploaded_file)
    st.image(image, caption='Imagen subida', use_column_width=True)
    
    if st.button('Generar PowerPoint'):
        with st.spinner('Leyendo texto...'):
            # 1. Extraer texto de la imagen
            # Nota: Streamlit Cloud ya tiene librerías de OCR instaladas
            texto_extraido = pytesseract.image_to_string(image, lang='spa')
            
            # 2. Crear el PowerPoint
            prs = Presentation()
            slide_layout = prs.slide_layouts[5] # Slide en blanco
            slide = prs.slides.add_slide(slide_layout)
            
            # Añadir el texto extraído
            textbox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(5))
            tf = textbox.text_frame
            tf.text = texto_extraido
            
            # 3. Guardar en memoria para descargar
            pptx_io = io.BytesIO()
            prs.save(pptx_io)
            pptx_io.seek(0)
            
            st.success("¡Listo! Tu PowerPoint está preparado.")
            st.download_button(
                label="📥 Descargar PowerPoint",
                data=pptx_io,
                file_name="notebook_editable.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
