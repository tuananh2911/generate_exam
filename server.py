import streamlit as st
import json
import random
from docx import Document
import io

# Đọc dữ liệu JSON
with open('output.json', 'r', encoding='utf-8') as file:
    data = json.load(file)

st.title("Tạo đề thi")

# Tạo các widget để chọn số câu hỏi cho mỗi bài và phần
selections = {}
for bai, content in data.items():
    st.header(bai)
    for phan, questions in content.items():
        key = f"{bai}_{phan}"
        num_questions = len(questions)
        selections[key] = st.number_input(f"Số câu hỏi cho {bai} {phan}", min_value=0, max_value=num_questions, value=0)

if st.button("Tạo đề thi"):
    # Tạo đề thi
    exam_questions = []
    for key, num_selected in selections.items():
        bai, phan = key.split('_')
        questions = data[bai][phan]
        selected_questions = random.sample(list(questions.items()), num_selected)
        exam_questions.extend([(bai, phan, q_num, q_text) for q_num, q_text in selected_questions])

    # Tạo file Word
    doc = Document()
    doc.add_heading('Đề thi', 0)

    for bai, phan, q_num, q_text in exam_questions:
        doc.add_paragraph(f"{bai} - {phan} - {q_num}: {q_text}")

    # Lưu file Word
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    # Tạo link tải xuống
    st.download_button(
        label="Tải xuống đề thi",
        data=buffer,
        file_name="de_thi.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )