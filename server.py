import streamlit as st
import json
import random
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENTATION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io
import zipfile

# Add a selection box for choosing the JSON file
json_option = st.selectbox(
    "Chọn đề cương",
    ("Đề cương lớp 10", "Đề cương lớp 11")
)

# Map the selection to the corresponding JSON file and headings
json_file_map = {
    "Đề cương lớp 10": {
        "file": "output_2.json",
        "headings": {
            "part1": "PHẦN I. Trắc nghiệm",
            "part2": "PHẦN II. Tự luận"
        }
    },
    "Đề cương lớp 11": {
        "file": "output.json",
        "headings": {
            "part1": "PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn. Thí sinh trả lời câu hỏi từ câu 1 đến câu 24. Mỗi câu hỏi thí sinh chỉ chọn một phương án.",
            "part2": "PHẦN II. Câu trắc nghiệm đúng, sai: Thí sinh trả lời từ câu 1 đến câu 4, trong mỗi ý a, b, c, d ở mỗi câu thí sinh chọn đúng hoặc sai."
        }
    }
}

selected_json_file = json_file_map[json_option]["file"]
selected_headings = json_file_map[json_option]["headings"]

# Read the selected JSON file
with open(selected_json_file, 'r', encoding='utf-8') as file:
    data = json.load(file)

st.title("Tạo đề thi")

# Create widgets to select the number of questions for each section
selections = {}
for bai, content in data.items():
    st.header(bai)
    for phan, questions in content.items():
        key = f"{bai}_{phan}"
        num_questions = len(questions)
        selections[key] = st.number_input(f"Số câu hỏi cho {bai} {phan}", min_value=0, max_value=num_questions, value=0)

# Add widget to select the number of exams
num_exams = st.number_input("Số lượng đề cần tạo", min_value=1, max_value=10, value=1)

def set_cell_border(cell, **kwargs):
    """
    Set cell's border
    Usage:
        set_cell_border(
            cell,
            top={"sz": 12, "val": "single", "color": "#FF0000", "space": "0"},
            bottom={"sz": 12, "color": "#00FF00", "val": "single"},
            start={"sz": 24, "val": "dashed", "shadow": "true"},
            end={"sz": 12, "val": "dashed"},
        )
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    # check for tag existence, if none found, then create one
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)

    # list over all available tags
    for edge in ('start', 'top', 'end', 'bottom', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = 'w:{}'.format(edge)

            # check for tag existence, if none found, then create one
            element = tcBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcBorders.append(element)

            # looks like order of attributes is important
            for key in ["sz", "val", "color", "space", "shadow"]:
                if key in edge_data:
                    element.set(qn('w:{}'.format(key)), str(edge_data[key]))

def format_question(doc, question_text):
    # Split question and answers
    parts = question_text.split('\n')
    question = parts[0]
    answers = parts[1:]

    # Add question to the same line as "Câu {i}:"
    p = doc.add_paragraph()
    p.add_run(question)
    for run in p.runs:
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)

    # Create table for answers
    if answers:
        table = doc.add_table(rows=len(answers), cols=1)
        table.allow_autofit = False
        table.width = Inches(6.5)  # Adjust table width

        for i, answer in enumerate(answers):
            cell = table.cell(i, 0)
            cell.text = answer.strip()

            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(12)

            # Remove cell border
            set_cell_border(
                cell,
                top={"sz": 0, "val": "none"},
                bottom={"sz": 0, "val": "none"},
                start={"sz": 0, "val": "none"},
                end={"sz": 0, "val": "none"},
            )

    # Add space after each question
    doc.add_paragraph()

if st.button("Tạo đề thi"):
    # Create a buffer to save the ZIP file
    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
        for exam_number in range(1, num_exams + 1):
            # Create exam
            exam_questions_part1 = []
            exam_questions_part2 = []
            for key, num_selected in selections.items():
                bai, phan = key.split('_')
                questions = data[bai][phan]
                selected_questions = random.sample(list(questions.items()), num_selected)
                if phan == "Phần 1":
                    exam_questions_part1.extend([(bai, phan, q_num, q_text) for q_num, q_text in selected_questions])
                elif phan == "Phần 2":
                    exam_questions_part2.extend([(bai, phan, q_num, q_text) for q_num, q_text in selected_questions])

            # Shuffle questions randomly in each part
            random.shuffle(exam_questions_part1)
            random.shuffle(exam_questions_part2)

            # Create Word file
            doc = Document()

            # Set default font
            style = doc.styles['Normal']
            style.font.name = 'Times New Roman'
            style.font.size = Pt(12)

            # Adjust page margins
            section = doc.sections[0]
            section.left_margin = Inches(0.5)
            section.right_margin = Inches(0.5)
            section.top_margin = Inches(0.5)
            section.bottom_margin = Inches(0.5)

            doc.add_heading(f'Đề thi - Mã đề {exam_number:03d}', 0)

            # Add Part 1 questions to the document with custom heading
            doc.add_heading(selected_headings["part1"], level=1)
            for i, (bai, phan, q_num, q_text) in enumerate(exam_questions_part1, 1):
                p = doc.add_paragraph()
                p.add_run(f"Câu {i}: ").bold = True
                p.add_run(q_text.split('\n')[0])  # Add question to the same line
                format_question(doc, '\n'.join(q_text.split('\n')[1:]))  # Only pass answers

            # Add Part 2 questions to the document with custom heading
            doc.add_heading(selected_headings["part2"], level=1)
            for i, (bai, phan, q_num, q_text) in enumerate(exam_questions_part2, 1):
                p = doc.add_paragraph()
                p.add_run(f"Câu {i}: ").bold = True
                p.add_run(q_text.split('\n')[0])  # Add question to the same line
                format_question(doc, '\n'.join(q_text.split('\n')[1:]))  # Only pass answers

            # Save Word file to buffer
            docx_buffer = io.BytesIO()
            doc.save(docx_buffer)
            docx_buffer.seek(0)

            # Add Word file to ZIP
            zip_file.writestr(f'de_thi_{exam_number:03d}.docx', docx_buffer.getvalue())

    # Prepare ZIP buffer for download
    zip_buffer.seek(0)

    # Create download link for ZIP file
    st.download_button(
        label="Tải xuống tất cả đề thi",
        data=zip_buffer,
        file_name="de_thi.zip",
        mime="application/zip"
    )