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

def get_cdf(cdf, database):
    for idx, cdf_row in cdf.iterrows():
    # 找到 df 中有包含 cdf_df 欄位文字的 row
        manu = str(cdf_row['Manufacturer/trademark']).strip()
        model = str(cdf_row['Type/model']).strip()
        matched = database[
            database['Manufacturer/trademark'].astype(str).str.strip().str.contains(manu, case=False, na=False, regex=False) &
            database['Type/model'].astype(str).str.strip().str.contains(model, case=False, na=False, regex=False)
        ]
        # 如果有找到，補資料
        if not matched.empty:
            for col in ['Object/part No.', 'Technical data', 'Standard', 'Mark(s) of conformity', 'website (UL)', 'VDE/TUV/ENEC']:
                cdf.at[idx, col] = matched.iloc[0][col]
    return cdf[columns]  


def run(cdf_path):
    # 讀取 Excel 檔案
    
    share_url = "https://z28856673-my.sharepoint.com/:x:/g/personal/itek_project_i-tek_com_tw/EThatAXi_QlGiUN0x9rNVaABawJbdCkryU5uRJyMVZazJg?e=JvWnha"

    # 多數 SharePoint 連結加上 download=1 可直下（若仍 403，請看下方 路徑B）
    download_url = share_url + ("&download=1" if "?" in share_url else "?download=1")

    headers = {"User-Agent": "Mozilla/5.0"}
    r = requests.get(download_url, headers=headers, allow_redirects=True, timeout=30)
    r.raise_for_status()  # 403/404 會在這裡丟錯

    df = pd.read_excel(BytesIO(r.content), sheet_name=0)  # 或指定 sheet_name

    cdf_df = pd.read_excel(cdf_path)
    cdf_df = get_cdf(cdf_df, df)
    output = BytesIO()
    cdf_df.to_excel(output, index=False)
    output.seek(0)
    return output
