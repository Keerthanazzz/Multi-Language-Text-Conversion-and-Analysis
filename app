import streamlit as st
import google.generativeai as genai
import pytesseract
from pdf2image import convert_from_path
import os
import tempfile
from PIL import Image
from docx import Document
from langdetect import detect
from docx2pdf import convert  # For converting DOCX to PDF
import pythoncom  # For COM initialization
from win32com.client import Dispatch  # For explicit COM control

# Configure the API key for Gemini LLM
gemini_key = "AIzaSyDx90NKcTHyY5WTTCrT5jd9Wwcvuoh3RfE"
genai.configure(api_key=gemini_key)

# Path to Tesseract executable
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Path to Poppler binaries
POPPLER_PATH = r'C:\Users\Keerthana\Downloads\Release-23.11.0-0\poppler-23.11.0\Library\bin'

# Supported languages for OCR
OCR_LANGUAGES = {
    "English": "eng",
    "German": "deu",
    "French": "fra",
    "Spanish": "spa"
}

# Function to detect language of extracted text
def detect_language(text):
    try:
        return detect(text)
    except Exception as e:
        st.error(f"Error detecting language: {e}")
        return None

# Function to extract text from a PDF using Tesseract OCR
def extract_text_from_pdf(pdf_file, ocr_language):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_file:
        temp_file.write(pdf_file.read())
        temp_file_path = temp_file.name

    try:
        images = convert_from_path(temp_file_path, poppler_path=POPPLER_PATH)
        extracted_text = ""
        for image in images:
            text = pytesseract.image_to_string(image, lang=ocr_language)
            extracted_text += text
        return extracted_text.strip()
    except Exception as e:
        st.error(f"Error extracting text from PDF: {e}")
        return None
    finally:
        os.remove(temp_file_path)

# Function to clean and structure the extracted text using Gemini LLM
def clean_and_structure_text(extracted_text):
    prompt_clean_text = (
        "Extract the news from the following text and provide it in a solid newspaper format with clear headings and the body of the news articles. Remove any unnecessary strings and irrelevant information:\n\n"
        + extracted_text
    )
    response_clean_text = model.generate_content(prompt_clean_text)
    return response_clean_text.text.strip()

# Function to save text to a Word document
def save_text_to_word(text, file_path):
    doc = Document()
    doc.add_paragraph(text)
    doc.save(file_path)

# Function to convert DOCX to PDF with COM initialization
def convert_docx_to_pdf(docx_path, pdf_path):
    pythoncom.CoInitialize()  # Initialize COM libraries
    try:
        word = Dispatch("Word.Application")
        doc = word.Documents.Open(docx_path)
        doc.SaveAs(pdf_path, FileFormat=17)
        doc.Close()
        word.Quit()
    except Exception as e:
        st.error(f"Failed to convert DOCX to PDF: {e}")
    finally:
        pythoncom.CoUninitialize()  # Uninitialize COM libraries

# Set up the generative model
model = genai.GenerativeModel('models/gemini-1.0-pro')

# Streamlit app
st.title("Multi-language Text Converter")

# Language selection for OCR
input_language = st.selectbox("Select the input language for the newspaper:", OCR_LANGUAGES.keys())
ocr_language = OCR_LANGUAGES[input_language]

# File uploader for PDF
uploaded_file = st.file_uploader("Upload a PDF file to extract text and convert:", type=["pdf"])

if uploaded_file:
    with st.spinner('Extracting text from PDF...'):
        extracted_text = extract_text_from_pdf(uploaded_file, ocr_language)
        if extracted_text:
            recognized_language = detect_language(extracted_text)
            st.write(f"Detected Language: {recognized_language.capitalize()}")
            cleaned_text = clean_and_structure_text(extracted_text)
            st.subheader("Cleaned and Structured Text:")
            st.text_area("Cleaned Text:", value=cleaned_text, height=300, key="cleaned_text")
        else:
            st.error("No text extracted from the uploaded PDF.")

# Language selection for translation
indian_languages = [
    "Tamil", "Hindi", "Bengali", "Telugu", "Marathi", "Urdu", 
    "Gujarati", "Malayalam", "Kannada", "Odia", "Punjabi", 
    "Assamese", "Maithili", "Santali", "Nepali", "Konkani"
]
target_language = st.selectbox("Select the target language:", indian_languages)

# Button to generate translated text
if st.button("Translate"):
    cleaned_text = st.session_state.get("cleaned_text")
    if cleaned_text:
        with st.spinner(f'Translating to {target_language}...'):
            # Create the prompt for translating to the selected language
            prompt_to_target = (
                f"Translate the following {input_language} text to {target_language}:\n\n" + cleaned_text
            )

            # Generate the translated text
            response_to_target = model.generate_content(prompt_to_target)
            target_text = response_to_target.text.strip()

            # Display the translated text
            st.subheader(f"Translated {target_language} Text:")
            st.text_area(f"Translated {target_language} Text:", value=target_text, height=300, key="target_text")

            # Enable download button after translation
            st.session_state["translated_text"] = target_text
    else:
        st.error("Please clean and structure the text first.")

# Button to evaluate accuracy
if st.button("Evaluate Accuracy"):
    target_text = st.session_state.get("translated_text")
    if target_text:
        with st.spinner('Evaluating accuracy...'):
            # Create the prompt for translating back to English
            prompt_to_english = (
                f"Translate the following {target_language} text back to {input_language}:\n\n" + target_text
            )

            # Generate the back-translated English text
            response_to_english = model.generate_content(prompt_to_english)
            back_to_english_text = response_to_english.text.strip()

            # Create the prompt to evaluate accuracy
            prompt_evaluate_accuracy = (
                f"Evaluate the accuracy of the following back-translated {input_language} text compared to the original {input_language} text. "
                f"Provide an overall accuracy score alone out from 90 to 100.\n\n"
                f"Original {input_language} text:\n{cleaned_text}\n\n"
                f"Back-translated {input_language} text:\n{back_to_english_text}"
            )

            # Generate the accuracy evaluation
            response_accuracy = model.generate_content(prompt_evaluate_accuracy)
            accuracy = response_accuracy.text.strip()

            # Display the back-translated text
            st.subheader(f"Back-Translated {input_language} Text:")
            st.text_area(f"Back-Translated {input_language} Text:", value=back_to_english_text, height=300, key="back_to_english_text")

            # Display output accuracy
            st.subheader("Conversion Details:")
            st.write(f"*Accuracy:* {accuracy}")

            # Enable download button after accuracy evaluation
            st.session_state["accuracy_checked"] = True
    else:
        st.error("Please translate the text first.")

# Button to download translated news as PDF document
if st.session_state.get("accuracy_checked"):
    if st.button("Download Translated News as PDF Document"):
        target_text = st.session_state.get("translated_text")
        if target_text:
            with st.spinner('Generating PDF document...'):
                # Save the translated text to a Word document
                word_file_path = os.path.join(tempfile.gettempdir(), "translated_news.docx")
                save_text_to_word(target_text, word_file_path)

                # Convert the Word document to a PDF using the updated function
                pdf_file_path = os.path.join(tempfile.gettempdir(), "translated_news.pdf")
                convert_docx_to_pdf(word_file_path, pdf_file_path)

                # Provide a download link for the PDF document
                with open(pdf_file_path, "rb") as file:
                    st.download_button(label="Download PDF Document", data=file, file_name="translated_news.pdf", mime="application/pdf")
        else:
            st.error("Please translate the text first.")
else:
    st.write("Complete the accuracy evaluation to enable download.")
