import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
import fitz  # PyMuPDF
import pytesseract
from PIL import Image
import io

st.set_page_config(page_title="PDF a PPTX Editable", page_icon="📊")

st.title("📊 Convertidor de PDF a PPTX Realmente Editable")
st.write("Esta versión usa OCR para intentar crear cuadros de texto que puedas editar.")

uploaded_file = st.file_uploader("Sube el PDF de NotebookLM", type="pdf")

if uploaded_file is not None:
    if st.button("Generar PowerPoint Editable"):
        with st.spinner("Procesando texto... Esto puede tardar un poco por página."):
            prs = Presentation()
            pdf_data = uploaded_file.read()
            doc = fitz.open(stream=pdf_data, filetype="pdf")
            
            for page in doc:
                # 1. Convertir página a imagen de alta calidad
                pix = page.get_pixmap(matrix=fitz.Matrix(3, 3))
                img = Image.open(io.BytesIO(pix.tobytes()))
                
                # 2. Crear diapositiva en blanco
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                
                # 3. OCR: Detectar texto y su posición
                data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT, lang='spa')
                
                # Agrupar palabras por líneas para que sea más fácil editar
                n_boxes = len(data['text'])
                for i in range(n_boxes):
                    if int(data['conf'][i]) > 60: # Solo si la confianza es alta
                        text = data['text'][i].strip()
                        if text:
                            # Convertir coordenadas de píxeles a pulgadas de PowerPoint
                            # (Ajuste aproximado de escala)
                            l = Inches(data['left'][i] / 300) 
                            t = Inches(data['top'][i] / 300)
                            w = Inches(data['width'][i] / 300)
                            h = Inches(data['height'][i] / 300)
                            
                            txBox = slide.shapes.add_textbox(l, t, w, h)
                            tf = txBox.text_frame
                            tf.text = text
                            # Intentar que el texto sea pequeño para que no se amontone
                            for paragraph in tf.paragraphs:
                                for run in paragraph.runs:
                                    run.font.size = Pt(12)

            pptx_io = io.BytesIO()
            prs.save(pptx_io)
            pptx_io.seek(0)
            
            st.success("¡Listo!")
            st.download_button(
                label="📥 Descargar PPTX Editable",
                data=pptx_io,
                file_name="presentacion_final.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
