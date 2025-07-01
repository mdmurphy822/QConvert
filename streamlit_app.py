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
        if re.match(r'^\d+\.', line):
            if current_question:
                questions.append(current_question)
            current_question = {"question": line, "options": [], "answer": ""}
        elif re.match(r'^[A-Da-d]\.', line):
            current_question["options"].append(line)
        elif line.lower().startswith("answer:"):
            current_question["answer"] = line
    if current_question:
        questions.append(current_question)

    vertical_rows = []
    for i, q in enumerate(questions, start=1):
        question_text = re.sub(r'^\d+\.\s*', '', q["question"]).strip('“”"')
        correct_answer = re.sub(r'^Answer:\s*', '', q["answer"], flags=re.IGNORECASE).strip('“”"')

        vertical_rows.append(["NewQuestion", "WR"])
        vertical_rows.append(["ID", f"Q_DocxImport_{i:02d}"])
        vertical_rows.append(["Title", f"Question {i}"])
        vertical_rows.append(["QuestionText", question_text])
        vertical_rows.append(["Points", 1])
        vertical_rows.append(["Difficulty", 5])
        vertical_rows.append(["Image", ""])
        vertical_rows.append(["InitialText", ""])
        vertical_rows.append(["AnswerKey", correct_answer])
        vertical_rows.append(["Hint", ""])
        vertical_rows.append(["Feedback", ""])
        vertical_rows.append([])

    df = pd.DataFrame(vertical_rows)
    csv_buffer = BytesIO()
    df.to_csv(csv_buffer, index=False, header=False)
    csv_buffer.seek(0)
    return csv_buffer

# Web UI
st.title("DOCX to Vertical CSV Quiz Converter")

uploaded_file = st.file_uploader("Upload your .docx quiz file", type="docx")

if uploaded_file:
    if st.button("Convert to CSV"):
        csv_output = convert_docx_to_vertical_csv(uploaded_file)
        st.success("✅ Conversion successful!")
        st.download_button(
            label="📥 Download CSV",
            data=csv_output,
            file_name="converted_quiz.csv",
            mime="text/csv"
        )
