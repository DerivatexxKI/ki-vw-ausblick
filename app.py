
import streamlit as st
import pandas as pd
import pdfplumber
from docx import Document
from io import BytesIO
import os
from openai import OpenAI

# OpenAI initialisieren
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

st.set_page_config(page_title="Volkswirtschaftlicher Ausblick – KI-Agent", page_icon="📊")
st.title("📊 KI-Agent für Volkswirtschaftlichen Ausblick")

# Datei-Upload
pdfs = st.file_uploader("📎 Lade Research-/Marktberichte als PDF hoch", type="pdf", accept_multiple_files=True)
excel = st.file_uploader("📈 Lade Excel-Datei mit Zinsdaten o. Ä.", type=["xls", "xlsx"])

# PDF-Inhalte extrahieren
def extract_text_from_pdfs(files):
    full_text = ""
    for pdf in files:
        with pdfplumber.open(pdf) as pdf_file:
            for page in pdf_file.pages:
                full_text += page.extract_text() + "\n"
    return full_text

# Excel-Inhalte extrahieren
def extract_excel_summary(file):
    try:
        df = pd.read_excel(file)
        return df.head().to_markdown()
    except Exception as e:
        return "Fehler beim Einlesen der Excel-Datei: " + str(e)

# Word-Dokument erstellen
def generate_word_report(prompt, title="Volkswirtschaftlicher Ausblick"):
    response = client.chat.completions.create(
        model="gpt-4",
        messages=[{"role": "user", "content": prompt}]
    )
    doc = Document()
    doc.add_heading(title, 0)
    doc.add_paragraph(response.choices[0].message.content)
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

if st.button("🧠 Analyse starten und Bericht generieren"):
    with st.spinner("Analysiere Inhalte und erstelle Bericht..."):
        pdf_text = extract_text_from_pdfs(pdfs) if pdfs else ""
        excel_text = extract_excel_summary(excel) if excel else ""

        gesamt_prompt = f"""Du bist ein erfahrener Volkswirt in einer Bank. Erstelle auf Basis folgender PDF-Auszüge und Excel-Daten einen professionellen volkswirtschaftlichen Ausblick für eine Bank:
---
PDF-Inhalte:
{pdf_text}
---
Excel-Daten (Vorschau):
{excel_text}
---
Gliedere den Text bitte wie folgt:
1. Aktuelle Wirtschaftslage
2. Zinsumfeld & Geldpolitik
3. Inflationsausblick
4. Risiken & geopolitische Spannungen
5. Mittelfristige Prognosen (inkl. Zinsen)

Sprache: sachlich, professionell, deutsch.""" 

        docx_file = generate_word_report(gesamt_prompt)
        st.success("Bericht erstellt!")
        st.download_button("📥 Word-Dokument herunterladen", data=docx_file, file_name="Volkswirtschaftlicher_Ausblick.docx")
