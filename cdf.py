import pandas as pd
from docx import Document

columns = [
    'Object/part No.', 'Manufacturer/trademark', 'Type/model',
    'Technical data', 'Standard', 'Mark(s) of conformity'
]
def save_cdf_to_word(df):
    # 建立一個 Word 文件
    doc = Document()
    # 加入表格：rows 為資料列數 + 標題列, cols 為欄位數
    table = doc.add_table(rows=1, cols=len(columns))
    table.style = 'Table Grid'

    # 寫入欄位標題
    hdr_cells = table.rows[0].cells
    for i, column in enumerate(columns):
        hdr_cells[i].text = str(column)

    for index, row in df.iterrows():
        row_cells = table.add_row().cells
        for i, col in enumerate(columns):
            value = row[col] if col in row else ""
            row_cells[i].text = str(value)

            # 將文字逐字處理，若是 ? 則變紅
            for char in value:
                run = para.add_run(char)
                if char == '?':
                    run.font.color.rgb = RGBColor(255, 0, 0)  # 紅色

    return doc


def run(cdf_path):
    # 讀取 Excel 檔案
    df = pd.read_excel('CDF_database_2025.03.27_Chris.xlsx')
    cdf_df = pd.read_excel(cdf_path)
    cdf_df = cdf_df.merge(
        df[['Object/part No.', 'Manufacturer/trademark', 'Type/model',
        'Technical data', 'Standard', 'Mark(s) of conformity', 'website (UL)', 'VDE/TUV/ENEC']],
        on=cdf_df.columns.to_list(),
        how='left'
        )

    doc = save_cdf_to_word(cdf_df)

    return doc

