import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor, Inches
import base64
import json
from openai import OpenAI
import os
from datetime import datetime

client = OpenAI(api_key='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')
chat
def response(fach, thema):
    response = client.chat.completions.create(
        model="gpt-3.5-turbo-1106",
        response_format={"type": "json_object"},
        messages=[
            {"role": "system", "content": "You are a helpful assistant designed to output JSON."},
            {"role": "user", "content": f"Erstelle ein Arbeitsblatt für das Fach {fach} zum {thema}. die ersten 5 fragen sollen verständnisfragen sein. die zweiten 2 fragen sollen multiple choise fragen sein(a), b), c), d)), und die letzte frage soll ein lückentext sein(c.a 4 sätze). es soll in diesem format sein : arbeitsblatt: ['Überschrift', 'Verständnisfrage', 'Verständnisfrage', 'Verständnisfrage', 'Verständnisfrage', 'Verständnisfrage 5', 'Multiple Frage choise a) antwort, b) antwort, c) antwort, d) antwort', 'Multiple choise', 'Lückentext']"}
        ]
    )
    while True:
        try:
            liste = json.loads(response.choices[0].message.content)["arbeitsblatt"]
            break
        except:
            pass
    return liste

def word(ab, fach):
    doc = Document()
    title = ab[0].replace(" ", "-").lower()

    sections = doc.sections
    for section in sections:
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)

    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri Light'
    font.size = Pt(12)
    font.color.rgb = RGBColor(40, 40, 40)
    
    aktueller_datum = datetime.now()
    datum = aktueller_datum.strftime("%d.%m.%y")
    header_style = doc.styles["Header"]
    header_paragraph = doc.sections[0].header.paragraphs[0]
    header_paragraph.text = f"{fach}\t\t{datum}"
    header_paragraph.style = header_style

    header_style = doc.styles["Heading1"] 
    header_paragraph = doc.add_paragraph()
    header_run = header_paragraph.add_run(ab[0])
    header_run.bold = True
    header_run.font.size = Pt(18) 

    for _ in range(2):
        doc.add_paragraph()

    for i, content in enumerate(ab[1:], start=1):
        subtitle_paragraph = doc.add_paragraph(f"{i}) {content}")
        subtitle_run = subtitle_paragraph.runs[0]
        subtitle_run.font.bold = True
        subtitle_run.font.size = Pt(14)
        subtitle_run.font.color.rgb = RGBColor(0, 28, 46)
        for _ in range(3):
            doc.add_paragraph()

    doc_name = f"{title}.docx"
    doc.save(doc_name)
    return title, doc_name

def fach():
    school_subjects = ["Mathematik 🔢","Deutsch 📚","Englisch 🇬🇧","Geschichte 📜","Geografie 🌍","Biologie 🌿","Chemie 🧪","Physik ⚙️","Informatik 💻", "Musik 🎵","Kunst 🎨","Sport 🏃‍♂️","Ethik 🤔","Religion ⛪","Politik 🗳️","Wirtschaft 💹","Philosophie 🤯", "Sozialkunde 👥","Psychologie 🧠","Sociology 👩‍👩‍👧‍👦","Fremdsprache 🗣️", "Latein 🏛️","Spanisch 🇪🇸", "Französisch 🇫🇷","Italienisch 🇮🇹","Russisch 🇷🇺","Chinesisch 🇨🇳","Japanisch 🇯🇵","Koreanisch 🇰🇷","Arabisch 🇸🇦","Medienkunde 📱",]

    subject_option = st.selectbox("Schulfach", ["Wähle ein Fach", "Eigenes Fach eingeben"] + school_subjects)

    if subject_option == "Wähle ein Fach":
        subject = None
    elif subject_option == "Eigenes Fach eingeben":
        subject = st.text_input("Eigenes Schulfach eingeben")[:-1]
        st.empty()
    else:
        subject = subject_option[:-1]

    return subject

st.set_page_config(
    page_title="CASE",
    page_icon="🏫",
)

st.title("CASE")

fach_selection = fach()
thema = st.text_input("Thema:")

if st.button("Arbeitsblatt erstellen"):
    ab = response(fach_selection, thema)
    title, doc_name = word(ab, fach_selection)
    st.success(f"Arbeitsblatt erfolgreich erstellt: {doc_name}")

    with open(doc_name, "rb") as f:
        data = f.read()
        b64 = base64.b64encode(data).decode('utf-8')
        href = f'<a href="data:application/octet-stream;base64,{b64}" download="{doc_name}">Dokument herunterladen</a>'
        st.markdown(href, unsafe_allow_html=True)

    os.remove(doc_name)
