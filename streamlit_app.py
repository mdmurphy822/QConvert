import re
import pandas as pd
import streamlit as st
from io import BytesIO
from docx import Document

def convert_docx_to_vertical_csv(docx_bytes):
    doc = Document(docx_bytes)
    lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]

    questions = []
    current_question = {}

    for line in lines:
        if re.match(r'^\d+\.', line):  # Start of new question
            if current_question:
                questions.append(current_question)
            current_question = {"question": line, "options": [], "answer": ""}
        elif re.match(r'^[A-Da-d]\.', line):  # MC option
            current_question["options"].append(line)
        elif line.lower().startswith("answer:"):  # Answer line
            current_question["answer"] = line
    if current_question:
        questions.append(current_question)

    vertical_rows = []
    for i, q in enumerate(questions, start=1):
        question_text = re.sub(r'^\d+\.\s*', '', q["question"]).strip('“”"')
        correct_line = re.sub(r'^Answer:\s*', '', q["answer"], flags=re.IGNORECASE).strip()
        correct_letter_match = re.match(r'^([A-Da-d])', correct_line)
        correct_letter = correct_letter_match.group(1).upper() if correct_letter_match else ""

        vertical_rows.append(["NewQuestion", "MC"])
        vertical_rows.append(["ID", f"Q_DocxImport_{i:02d}"])
        vertical_rows.append(["Title", f"Question {i}"])
        vertical_rows.append(["QuestionText", question_text])
        vertical_rows.append(["Points", 1])
        vertical_rows.append(["Difficulty", 5])
        vertical_rows.append(["Image", ""])

        for option in q["options"]:
            match = re.match(r'^([A-Da-d])\.\s*(.*)', option)
            if not match:
                continue
            letter, text = match.groups()
            score = 100 if letter.upper() == correct_letter else 0
            vertical_rows.append(["Option", score, text])

        vertical_rows.append(["Hint", ""])      # Optional row
        vertical_rows.append(["Feedback", ""])  # Optional row
        vertical_rows.append([])                # Blank line = question separator

    df = pd.DataFrame(vertical_rows)
    csv_buffer = BytesIO()
    df.to_csv(csv_buffer, index=False, header=False, encoding="utf-8")
    csv_buffer.seek(0)
    return csv_buffer

# Streamlit Web UI
st.title("📄 DOCX to Brightspace MC CSV Converter")

uploaded_file = st.file_uploader("Upload your quiz .docx file", type="docx")

if uploaded_file:
    if st.button("Convert to Brightspace CSV"):
        csv_output = convert_docx_to_vertical_csv(uploaded_file)
        st.success("✅ Conversion complete!")
        st.download_button(
            label="📥 Download CSV File",
            data=csv_output,
            file_name="converted_quiz.csv",
            mime="text/csv"
        )
