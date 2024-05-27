from docx import Document
from docx.oxml.table import CT_Tbl
from docx.table import Table


def fill_table_after_heading(doc, key, value):
    found_target = False

    for element in doc.element.body:
        if isinstance(element, CT_Tbl) and found_target:
            table = Table(element, doc)
            for row in table.rows:
                for i, cell in enumerate(row.cells):
                    if cell.text.strip() == "模块名称":
                        if i + 1 < len(row.cells):
                            row.cells[i + 1].text = key
                    elif cell.text.strip() == "功能描述":
                        if i + 1 < len(row.cells):
                            row.cells[i + 1].text = value
                    elif cell.text.strip() == "修改说明":
                        if i + 1 < len(row.cells):
                            row.cells[i + 1].text = value
            break  # 找到并处理第一个表格后退出循环
        elif element.tag.endswith('p'):
            paragraph = element
            if paragraph.text.strip() == key:
                found_target = True


def process_document(doc_path, data):
    doc = Document(doc_path)

    for key, value in data.items():
        fill_table_after_heading(doc, key, value)

    doc.save(doc_path)  # 保存修改后的文档







