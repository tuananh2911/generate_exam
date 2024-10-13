import re

import docx
import json

def read_docx(file_path):
    doc = docx.Document(file_path)
    full_text = []
    for para in doc.paragraphs:
        full_text.append(para.text)
    return '\n'.join(full_text)

# Đường dẫn đến file docx
file_path = 'de_cuong.docx'

# Đọc nội dung file
content = read_docx(file_path)

# Tách nội dung theo "Bài:"
bai_list = content.split('BÀI ')
obj = {}
for bai_index, bai in enumerate(bai_list):  # Bắt đầu từ phần tử thứ 2 và đánh số từ 1
    bai_obj = {}
    phan_list = bai.split('Phần ')
    for phan_index, phan in enumerate(phan_list[1:]):  # Bắt đầu từ phần tử thứ 2 và đánh số từ 1
        cau_list = re.split(r'Câu \d+.', phan)
        print(cau_list[0])
        phan_obj = {}
        for cau_index, cau in enumerate(cau_list[1:], 1):  # Bắt đầu từ phần tử thứ 2 và đánh số từ 1
            phan_obj[f'Câu {cau_index}'] = cau.strip()
        bai_obj[f'Phần {phan_index+1}'] = phan_obj
    obj[f'BÀI {bai_index+1}'] = bai_obj

# Chuyển đổi object thành JSON
json_output = json.dumps(obj, ensure_ascii=False, indent=2)


# Lưu JSON vào file
with open('output.json', 'w', encoding='utf-8') as f:
    json.dump(obj, f, ensure_ascii=False, indent=2)