import pandas as pd
from docx import Document
from docx.shared import RGBColor
import re
from docx.shared import Pt
import requests
from io import BytesIO


columns = [
    'Object/part No.', 'Manufacturer/trademark', 'Type/model',
    'Technical data', 'Standard', 'Mark(s) of conformity', 'website (UL)', 'VDE/TUV/ENEC'
]

def comp_translation(comp, comp_df):
    result = comp_df[comp_df['english'] == comp]
    if result.empty:
        return comp
    else:
        return result['chinese'].values[0]


def save_cdf_to_word(df):
     # 讀取 Excel 檔案
    
    share_url = "https://z28856673-my.sharepoint.com/:x:/g/personal/itek_project_i-tek_com_tw/EV_kkQot_hZNsD8LFYQLfqoBWvR5p28e8_7yvoQXeVHtkg?e=SgUjeV"

    # 多數 SharePoint 連結加上 download=1 
    download_url = share_url + ("&download=1" if "?" in share_url else "?download=1")

    headers = {"User-Agent": "Mozilla/5.0"}
    r = requests.get(download_url, headers=headers, allow_redirects=True, timeout=30)
    r.raise_for_status()  # 403/404 會在這裡丟錯

    comp_df = pd.read_excel(BytesIO(r.content), sheet_name=0)  # 或指定 sheet_name
    
    
    comp_df = pd.read_excel('component_translate.xlsx')
    doc = Document()
    # 加入表格：rows 為資料列數 + 標題列, cols 為欄位數
    table = doc.add_table(rows=1, cols=len(columns))
    table.style = 'Table Grid'

    # 設定欄寬（單位: pt，1英吋=72pt，可依需求調整）
    col_widths = [100, 100, 100, 100, 100, 100, 100, 100]
    for i, width in enumerate(col_widths):
        for cell in table.columns[i].cells:
            cell.width = Pt(width)

    # 寫入欄位標題
    hdr_cells = table.rows[0].cells
    for i, column in enumerate(columns):
        hdr_cells[i].text = str(column)
        p = hdr_cells[i].paragraphs[0]
        p.alignment = 0  # 置左
        for run in p.runs:
            run.font.size = Pt(11)
            run.bold = True


    for index, row in df.iterrows():
        row_cells = table.add_row().cells
        for i, col in enumerate(columns):
            value = row[col] if col in row else ""
            row_cells[i].text = str(value)
            text = str(value)

            if i == 0:  # 第一列是元件名稱
                text = comp_translation(text, comp_df)
            # 去除 UL 的部分
            segs = [s.strip() for s in text.split(",")]
            segs = [s for s in segs if s and "UL" not in s.upper()]
            text = ", ".join(segs)

            parts = re.split(r'([?])', text)  # 保留分隔符
            p = row_cells[i].paragraphs[0]              # 你的目標段落
            p.alignment = 0  # 置左
            # 先清空原有 runs（保留段落樣式）
            for r in p.runs:
                p._p.remove(r._r)

            for part in parts:
                r = p.add_run(part)
                r.font.size = Pt(10)
                if text in ('?', '？') and i <= 5:  # 只對前6列加紅色
                    r.font.color.rgb = RGBColor(255, 0, 0)

    return doc


def run(cdf_path):
    # 讀取 Excel 檔案
    
    share_url = "https://z28856673-my.sharepoint.com/:x:/g/personal/itek_project_i-tek_com_tw/EThatAXi_QlGiUN0x9rNVaABawJbdCkryU5uRJyMVZazJg?e=JvWnha"

    # 多數 SharePoint 連結加上 download=1 可直下（若仍 403，請看下方 路徑B）
    download_url = share_url + ("&download=1" if "?" in share_url else "?download=1")

    headers = {"User-Agent": "Mozilla/5.0"}
    r = requests.get(download_url, headers=headers, allow_redirects=True, timeout=30)
    r.raise_for_status()  # 403/404 會在這裡丟錯

    df = pd.read_excel(BytesIO(r.content), sheet_name=0)  # 或指定 sheet_name


    # df = pd.read_excel(url, sheet_name='CDF_database_2025.03.27_Chris')
    
    cdf_df = pd.read_excel(cdf_path)
    cdf_df = cdf_df.merge(
        df[['Object/part No.', 'Manufacturer/trademark', 'Type/model',
        'Technical data', 'Standard', 'Mark(s) of conformity', 'website (UL)', 'VDE/TUV/ENEC']],
        on=cdf_df.columns.to_list(),
        how='left'
        )
    doc = save_cdf_to_word(cdf_df)

    return doc
