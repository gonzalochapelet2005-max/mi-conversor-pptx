import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
import fitz  # PyMuPDF
import io

st.set_page_config(page_title="PDF a PowerPoint Editable", page_icon="📊")

st.title("📊 Convertidor de PDF (NotebookLM) a PPTX")
st.write("Sube el PDF que te dio NotebookLM y lo convertiré en un PowerPoint editable.")

uploaded_file = st.file_uploader("Elige tu archivo PDF", type="pdf")

if uploaded_file is not None:
    if st.button("Generar PowerPoint"):
        with st.spinner("Convirtiendo páginas..."):
            # Crear presentación
            prs = Presentation()
            
            # Abrir el PDF desde la memoria
            pdf_data = uploaded_file.read()
            doc = fitz.open(stream=pdf_data, filetype="pdf")
            
            for page in doc:
                # 1. Extraer imagen de la página para el fondo
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                img_data = pix.tobytes("png")
                
                # 2. Crear slide
                slide_layout = prs.slide_layouts[6] # Blanco
                slide = prs.slides.add_slide(slide_layout)
                
                # 3. Poner la imagen de fondo (no editable)
                img_stream = io.BytesIO(img_data)
                slide.shapes.add_picture(img_stream, 0, 0, width=prs.slide_width, height=prs.slide_height)
                
                # 4. Extraer texto y ponerlo encima (editable)
                text_blocks = page.get_text("blocks")
                for b in text_blocks:
                    left = Inches(b[0] / 72)
                    top = Inches(b[1] / 72)
                    width = Inches((b[2] - b[0]) / 72)
                    height = Inches((b[3] - b[1]) / 72)
                    
                    txBox = slide.shapes.add_textbox(left, top, width, height)
                    tf = txBox.text_frame
                    tf.text = b[4]
                    tf.word_wrap = True

            # Guardar
            pptx_io = io.BytesIO()
            prs.save(pptx_io)
            pptx_io.seek(0)
            
            st.success("¡Listo!")
            st.download_button(
                label="📥 Descargar PowerPoint Editable",
                data=pptx_io,
                file_name="presentacion_editable.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
