import re
from docx import Document
import pandas as pd
import streamlit as st

class DocTableResovler:
    @classmethod
    def extract_docs_tables(cls, doc, required_keys=None):
        tables = doc.tables
        data_dict = {}
        for table in tables:
            table_data = cls.extract_table_data(table, required_keys)
            for key, value in table_data.items():
                if key in data_dict:
                    data_dict[key] += f"\n{value}"
                else:
                    data_dict[key] = value
        return data_dict

    @classmethod
    def extract_table_data(cls, table, required_keys=None):
        try:
            rows = table.rows
            cols = len(rows[0].cells) if rows else 0
        except IndexError:
            return {}

        data_dict = {}
        for row in rows:
            row_data = cls.extract_table_row(row)
            for key, value in row_data.items():
                if key in data_dict:
                    data_dict[key] += f"\n{value}"
                else:
                    data_dict[key] = value

        return (
            data_dict
            if not required_keys
            else {
                k: v
                for k, v in data_dict.items()
                if any(search_str in k for search_str in required_keys)
            }
        )

    @classmethod
    def extract_table_row(cls, row):
        row_data_dict = {}
        cells = row.cells
        cell_num = len(cells)
        if cell_num <= 1:
            return row_data_dict
        question = None
        i = 0
        while i < cell_num:
            if not question:
                question = cells[i].text
                i += 1
                continue
            if question == cells[i].text:
                i += 1
                continue
            row_data_dict[question.strip().replace("\n", "")] = (
                cells[i].text.strip().replace("\n", "")
            )
            question = None
            while i < cell_num - 1 and cells[i + 1].text == cells[i].text:
                i += 1
            i += 1
        return row_data_dict


def extract_goal_codes(text):
    # 使用正则表达式提取所有类似 (A1), (B2) 的字母数字组合
    pattern = r"[A-D][1-9]"  # 匹配 (A1), (B2), (C3) 等格式
    matches = re.findall(pattern, text)
    print(matches)
    # 去重并按字母升序排序
    unique_matches = sorted(set(matches), key=lambda x: (x[0], int(x[1:])))

    return unique_matches


# ======================== Streamlit Web 应用 ======================== #
st.title("📄 Word 课程信息解析")

# 1️⃣ 让用户上传 Word 文件
uploaded_file = st.file_uploader("请上传课程教学大纲 (Word 文档)", type=["docx"])

# 2️⃣ 用户输入需要匹配的字符串（字段）
required_keys_input = st.text_area(
    "请输入所需字段（字符串），用英文逗号 `,` 分隔",
    "课程代码,课程名称,学时,学分,课程目标 (Course Object) ",
)

if uploaded_file:
    # 解析用户输入的 required_keys
    required_keys = [
        key.strip() for key in required_keys_input.split(",") if key.strip()
    ]

    # 3️⃣ 读取 Word 文档
    doc = Document(uploaded_file)

    # 4️⃣ 解析表格数据
    parsed_data = DocTableResovler.extract_docs_tables(doc, required_keys)

    # 5️⃣ 转换为 DataFrame
    df = pd.DataFrame(parsed_data.items(), columns=["字段", "值"])

    # 6️⃣ 显示表格
    st.subheader("解析结果：")
    st.table(df)  # 显示静态表格

    # 7️⃣ 如果有 课程目标，提取目标代码
    course_object_keys = ["课程目标 (Course Object)", "*课程目标 (Course Object)"]
    for key in course_object_keys:
        if key in parsed_data:
            goal_codes = extract_goal_codes(parsed_data[key])
            st.write(f"**提取的课程目标代码 ({key})**：")
            st.write(", ".join(goal_codes))
        else:
            st.warning("没有匹配的数据，请检查您的字段！")

