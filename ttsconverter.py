import streamlit as st
from docx import Document
import pyttsx3
import os
import shutil
import openai

# Set your OpenAI API key here
openai.api_key = "OPEN_API_KEY"

# Ensure necessary directories exist
os.makedirs('uploads', exist_ok=True)
os.makedirs('output', exist_ok=True)

def extract_text_from_docx(docx_file):
    doc = Document(docx_file)
    full_text = []
    for para in doc.paragraphs:
        full_text.append(para.text)
    return '\n'.join(full_text)

def convert_text_to_speech(text, output_path):
    try:
        engine = pyttsx3.init()
        engine.save_to_file(text, output_path)
        engine.runAndWait()
    except Exception as e:
        return str(e)
    return None

def clear_directory(dir_path):
    if os.path.exists(dir_path):
        shutil.rmtree(dir_path)
    os.makedirs(dir_path, exist_ok=True)

def latex_to_readable(latex_code):
    combined_prompt = f"""
    You are an intelligent assistant. Your task is to convert LaTeX code in the given text to plain English and provide summaries for tables and images. Follow these specific instructions:

    1. Keep the plain text exactly as it is. Do not provide any overview or summary of the entire text.
    2. Convert any LaTeX code into a human-readable format. For example, $\\alpha$ should be read as "alpha", and $ax^2+bx+c=0$ should be read as "a x squared plus b x plus c equals zero".
    3. For tables, provide a summarized explanation of the concept covered in the table without mentioning the formatting. For example, if a table shows a logical representation of inputs and outputs, explain the logical relationship in plain English.
    4. For images, mention that there is an image at that position, but do not narrate the source or provide additional context about the image.

    Text to convert:
    {latex_code}
    """
    
    try:
        response = openai.Completion.create(
            engine="davinci-codex",  # Use the Codex engine for code understanding
            prompt=combined_prompt,
            max_tokens=150  # Adjust max_tokens as per your requirement
        )
        human_readable_text = response['choices'][0]['text'].strip()
    except Exception as e:
        return f"Error: {e}"
    
    return human_readable_text

# Streamlit app
st.title("LaTeX to Speech Converter")

st.write("""Upload a DOCX file with LaTeX content, and get an MP3 narration.""")

uploaded_file = st.file_uploader("Choose a DOCX file", type="docx")

if uploaded_file is not None:
    clear_directory('uploads')
    clear_directory('output')
    
    # Save uploaded file
    file_path = os.path.join('uploads', uploaded_file.name)
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    text = extract_text_from_docx(file_path)

    # Use OpenAI API to convert LaTeX to human-readable text
    try:
        with st.spinner("Converting LaTeX to human-readable text..."):
            converted_text = latex_to_readable(text)
        st.success
