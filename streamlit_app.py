import streamlit as st
import pandas as pd
import re
from docx import Document
from io import BytesIO

def embedded_block_parser_from_docx(docx_file):
    doc = Document(docx_file)
    lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    vertical_rows = []

    for line in lines:
        parts = re.split(r'(?=[A-Da-d]\.)', line)
        if len(parts) < 5:
            continue  # skip malformed lines

        question_raw = parts[0]
        options_raw = parts[1:5]

        # Extract clean question
        question_text = re.sub(r'^\d+\.\s*', '', question_raw).strip()

        # Extract and remove embedded answer from last option
        answer_match = re.search(r'Answer:\s*([A-Da-d])', options_raw[3], re.IGNORECASE)
        answer_letter = answer_match.group(1).upper() if answer_match else ""
        options_raw[3] = re.sub(r'Answer:\s*[A-Da-d]', '', options_raw[3], flags=re.IGNORECASE).strip()

        vertical_rows.append(["NewQuestion", "MC"])
        vertical_rows.append(["ID", ""])
        vertical_rows.append(["Title", ""])
        vertical_rows.append(["QuestionText", question_text])
        vertical_rows.append(["Points", 1])
        vertical_rows.append(["Difficulty", 1])
        vertical_rows.append(["Image", ""])

        for opt in options_raw:
            match = re.match(r'^([A-Da-d])\.\s*(.*)', opt.strip())
            if not match:
                continue
            letter, text = match.groups()
            weight = "100" if letter.upper() == answer_letter else "0"
            vertical_rows.append(["Option", weight, text.strip()])

        vertical_rows.append(["Hint", ""])
        vertical_rows.append(["Feedback", ""])
        vertical_rows.append([])

    return pd.DataFrame(vertical_rows)

# --- Streamlit App ---
st.set_page_config(page_title="DOCX to Brightspace Quiz", layout="wide")
st.title("📄 DOCX to Brightspace Quiz Converter")

uploaded_file = st.file_uploader("Upload your .docx quiz file", type=["docx"])

if uploaded_file:
    try:
        df = embedded_block_parser_from_docx(uploaded_file)

        st.success("✅ File successfully parsed!")
        st.dataframe(df, use_container_width=True)

        # Export to CSV for download
        csv_bytes = df.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="📥 Download Brightspace CSV",
            data=csv_bytes,
            file_name="converted_quiz.csv",
            mime="text/csv"
        )

    except Exception as e:
        st.error(f"⚠️ Error processing file: {str(e)}")
else:
    st.info("Please upload a .docx file to begin.")
