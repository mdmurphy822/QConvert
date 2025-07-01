import re
import pandas as pd
from docx import Document
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import os

def convert_docx_to_vertical_csv(docx_path, output_csv_path):
    doc = Document(docx_path)
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
    df.to_csv(output_csv_path, index=False, header=False)
    print(f"\n✅ Successfully saved to: {output_csv_path}")

def main():
    # Open file explorer to select .docx file
    Tk().withdraw()  # Hide root window
    print("📂 Please select the .docx file containing your quiz...")
    docx_path = askopenfilename(filetypes=[("Word Documents", "*.docx")])
    
    if not docx_path:
        print("❌ No file selected. Exiting.")
        return
    
    output_csv_path = os.path.splitext(docx_path)[0] + "_VerticalFormat.csv"
    convert_docx_to_vertical_csv(docx_path, output_csv_path)

if __name__ == "__main__":
    main()
