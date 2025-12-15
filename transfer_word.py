import pandas as pd
from docx import Document
from copy import deepcopy

columns = [
    'Object/part No.', 'Manufacturer/trademark', 'Type/model',
    'Technical data', 'Standard', 'Mark(s) of conformity'
]

def style_setting(doc, bsmi_on):
    if bsmi_on:
        style = doc.styles['Normal']
        style.font.name = 'Arial'
        style._element.rPr.rFonts.set(qn('w:eastAsia'), "標楷體")
        style.font.size = Pt(11)
    else:
        style = doc.styles['Normal']
        style.font.name = 'Arial'
        style.font.size = Pt(10)

def WriteInDataSheet(doc, cdf_path, bsmi_on):

    cdf_df = pd.read_excel(cdf_path)
    cdf_df = cdf_df[columns]
    table = doc.tables[1]  # 假設只處理第一個表格
    
    for r_idx, row in enumerate(table.rows):
        texts = [cell.text.strip().lower() for cell in row.cells]
        if "Object/part No.".lower() in texts:  # 找到表頭 Location
            start_row = r_idx

    target_tr = table.rows[start_row]._tr
    insert_pos = list(table._tbl).index(target_tr) + 2  # 插入到這一列之前
    for idx, row in cdf_df.iterrows():
        new_tr = deepcopy(table.rows[start_row + 1]._tr)
        table._tbl.insert(insert_pos, new_tr)
        new_row_idx = list(table._tbl.tr_lst).index(new_tr)
        inserted_row = table.rows[new_row_idx]
        # 依 columns 順序填入每個 cell
        for col_idx, col in enumerate(columns):
            inserted_row.cells[col_idx].text = str(row[col])
        insert_pos += 1  # 下一個插在後面

    table._tbl.remove(table.rows[start_row + 1]._tr)  # 刪除原有的空白行
    style_setting(doc, bsmi_on)
    return doc
