import os
import streamlit as st
from docx import Document
import pythoncom
from docx2pdf import convert
from pdf2image import convert_from_path
import zipfile
from fpdf import FPDF
import pyttsx3
import tempfile
from PyPDF2 import PdfReader


st.set_page_config(page_title='PDFOT', page_icon='pdf.png', layout="centered", initial_sidebar_state="auto", menu_items=None)

hide_streamlit_style = """
    <style>
    footer {visibility: hidden;}
    </style>
    """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

def pdf2docx():
    def convert_pdf_to_doc(pdf_path, doc_path):
        pdf_reader = PdfReader(pdf_path)

        doc = Document()

        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            text = page.extract_text()
            doc.add_paragraph(text)

        doc.save(doc_path)

    def main():
        st.markdown(
            "<center><h1 style='font-family: Comic Sans MS; font-weight: 300; font-size: 32px;'>PDF to DOCX "
            "Converter</h1></center>",
            unsafe_allow_html=True)
        st.markdown(
            "<center><h1 style='font-family: Comic Sans MS; font-weight: 300; font-size: 12px;'>Upload a PDF file and "
            "convert it to DOCX.</h1></center>",
            unsafe_allow_html=True)

        uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")

        if uploaded_file is not None:
            with st.spinner('Converting PDF to DOCX...'):
                with open("temp.pdf", "wb") as file:
                    file.write(uploaded_file.getbuffer())

                convert_pdf_to_doc("temp.pdf", "converted.docx")

                st.success("Conversion completed!")

                st.download_button(
                    label="Download Converted DOC",
                    data=open("converted.docx", "rb").read(),
                    file_name="converted.docx"
                )
                os.remove("temp.pdf")
                os.remove("converted.docx")

    if __name__ == '__main__':
        main()


def docx2pdf():
    def convert_docx_to_pdf(docx_path, pdf_path):
        pythoncom.CoInitialize()
        convert(docx_path, pdf_path)

    def main():
        st.markdown(
            "<center><h1 style='font-family: Comic Sans MS; font-weight: 300; font-size: 32px;'>DOCX to PDF "
            "Converter</h1></center>",
            unsafe_allow_html=True)
        st.markdown(
            "<center><h1 style='font-family: Comic Sans MS; font-weight: 300; font-size: 12px;'>Upload a DOCX file "
            "and convert it to PDF.</h1></center>",
            unsafe_allow_html=True)

        uploaded_file = st.file_uploader("Choose a DOCX file", type="docx")
        conversion_completed = False

        if uploaded_file is not None:
            with st.spinner('Converting DOCX to PDF...'):
                temp_path = os.path.join(os.getcwd(), "temp.docx")
                with open(temp_path, "wb") as file:
                    file.write(uploaded_file.getbuffer())
                convert_docx_to_pdf(temp_path, "converted.pdf")

                st.success("Conversion completed!")
                conversion_completed = True
                os.remove(temp_path)

        if conversion_completed:
            with open("converted.pdf", "rb") as file:
                pdf_data = file.read()
            st.download_button(
                label="Download Converted PDF",
                data=pdf_data,
                file_name="converted.pdf"
            )

    if __name__ == '__main__':
        main()


def pdf2png():
    def convert_pdf_to_png(pdf_path, output_dir):
        images = convert_from_path(pdf_path)
        for i, image in enumerate(images):
            image_path = os.path.join(output_dir, f"page_{i + 1}.png")
            image.save(image_path, "PNG")

    def main():
        st.markdown(
            "<center><h1 style='font-family: Comic Sans MS; font-weight: 300; font-size: 32px;'>PDF to PNG "
            "Converter</h1></center>",
            unsafe_allow_html=True)
        st.markdown(
            "<center><h1 style='font-family: Comic Sans MS; font-weight: 300; font-size: 12px;'>Upload a PDF file and "
            "convert it to PNG images.</h1></center>",
            unsafe_allow_html=True)

        uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")

        if uploaded_file is not None:
            with st.spinner('Converting PDF to PNG...'):
                with open("temp.pdf", "wb") as file:
                    file.write(uploaded_file.getbuffer())
                output_dir = "output"
                os.makedirs(output_dir, exist_ok=True)
                convert_pdf_to_png("temp.pdf", output_dir)

                st.success("Conversion completed!")
                with tempfile.NamedTemporaryFile(delete=False) as temp_zip:
                    with zipfile.ZipFile(temp_zip, "w") as zip_file:
                        for file_name in os.listdir(output_dir):
                            file_path = os.path.join(output_dir, file_name)
                            zip_file.write(file_path, file_name)
                    st.download_button(
                        label="Download Converted Images",
                        data=open(temp_zip.name, "rb").read(),
                        file_name="converted_images.zip",
                        mime="application/zip"
                    )

                os.remove("temp.pdf")
                for file_name in os.listdir(output_dir):
                    file_path = os.path.join(output_dir, file_name)
                    os.remove(file_path)
                os.rmdir(output_dir)

    if __name__ == '__main__':
        main()


def png2pdf():
    def convert_images_to_pdf(image_paths, output_path):
        pdf = FPDF()

        for image_path in image_paths:
            pdf.add_page()
            pdf.image(image_path, x=0, y=0, w=210, h=297)  # Adjust width (w) and height (h) as needed

        pdf.output(output_path, "F")

    def main():
        st.markdown(
            "<center><h1 style='font-family: Comic Sans MS; font-weight: 300; font-size: 32px;'>Image to PDF "
            "Converter</h1></center>",
            unsafe_allow_html=True)
        st.markdown(
            "<center><h1 style='font-family: Comic Sans MS; font-weight: 300; font-size: 12px;'>Upload PNG or JPG "
            "images and convert them to a single PDF.</h1></center>",
            unsafe_allow_html=True)

        uploaded_files = st.file_uploader("Choose PNG or JPG images", type=["png", "jpg"], accept_multiple_files=True)

        if uploaded_files:
            with st.spinner('Converting Images to PDF...'):
                temp_dir = "temp_images"
                os.makedirs(temp_dir, exist_ok=True)

                image_paths = []
                for i, uploaded_file in enumerate(uploaded_files):
                    image_path = os.path.join(temp_dir, f"image_{i + 1}.{uploaded_file.name.split('.')[-1]}")
                    with open(image_path, "wb") as file:
                        file.write(uploaded_file.getbuffer())
                    image_paths.append(image_path)

                output_path = "converted.pdf"
                convert_images_to_pdf(image_paths, output_path)

                st.success("Conversion completed!")
                st.download_button(
                    label="Download Converted PDF",
                    data=open(output_path, "rb").read(),
                    file_name="converted.pdf"
                )

                for image_path in image_paths:
                    os.remove(image_path)
                os.rmdir(temp_dir)
                os.remove(output_path)

    if __name__ == '__main__':
        main()


def pdf2audio():
    def convert_pdf_to_audio(file, voice):
        pdf = PdfReader(file)
        text = ""
        for page in pdf.pages:
            text += page.extract_text()

        engine = pyttsx3.init()
        voices = engine.getProperty("voices")
        if voice >= len(voices):
            voice = 0
        engine.setProperty("voice", voices[voice].id)
        output_file = tempfile.NamedTemporaryFile(delete=False, suffix=".mp3").name
        engine.save_to_file(text, output_file)
        engine.runAndWait()

        return output_file

    def main():
        st.markdown(
            "<center><h1 style='font-family: Comic Sans MS; font-weight: 300; font-size: 32px;'>PDF to Audio "
            "Converter</h1></center>",
            unsafe_allow_html=True)
        file = st.file_uploader("Upload a PDF file", type="pdf")
        engine = pyttsx3.init()
        voices = engine.getProperty("voices")
        voice_options = [f"{i}. {voices[i].name}" for i in range(len(voices))]
        default_voice = 0
        voice = st.selectbox("Select a voice", voice_options, index=default_voice)

        if st.button("Convert to Audio"):
            if file is not None:
                output_file = convert_pdf_to_audio(file, int(voice.split(".")[0]))
                st.success("Conversion completed. Click below to download the audio file.")
                st.audio(output_file, format="audio/mp3")
                with open(output_file, "rb") as f:
                    audio_bytes = f.read()
                st.download_button(
                    label="Download Audio File",
                    data=audio_bytes,
                    file_name="converted_audio.mp3",
                    mime="audio/mp3",
                )
                os.remove(output_file)

    if __name__ == "__main__":
        main()



def main():
    st.sidebar.markdown("""
            <style>
                .sidebar-text {
                    text-align: center;
                    font-size: 32px;
                    font-weight: bold;
                    font-family: Comic Sans MS;
                }
            </style>
            <p class="sidebar-text">PDF</p>
        """, unsafe_allow_html=True)
    st.sidebar.image(
        "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcR5E66uvCKRTpFG7sxxGiM9YqG77gp8mTK8V3YNkSUfI94WBktk0Te8"
        "-LkNCtSoNyG33RY&usqp=CAU",
        use_column_width=True)
    selected_sidebar = st.sidebar.radio("Please Select One", ["PDF to DOCX", "DOCX to PDF","PDF to PNG", "Image to PDF","PDF to AUDIO"])

    if selected_sidebar == "PDF to DOCX":
        pdf2docx()
    elif selected_sidebar == "DOCX to PDF":
        docx2pdf()
    elif selected_sidebar == "PDF to PNG":
        pdf2png()
    elif selected_sidebar == "Image to PDF":
        png2pdf()
    elif selected_sidebar == "PDF to AUDIO":
        pdf2audio()



if __name__ == "__main__":
    main()