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

def clean_str(x):
    if pd.isna(x):
        return ""
    s = str(x).strip()
    return "" if s.lower() in {"nan", "none"} else s

def check_y_capacitor(obj, tech_data):
    print(obj, tech_data)
    if obj in tech_data:
        return True
    else:
        return False

def get_cdf(cdf, database):
    output = pd.DataFrame(columns=columns)

    # 預先把 database 要比對的欄位正規化（加速與一致性）
    db_model = database['Type/model'].astype(str).str.strip().str.lower()
    db_manu  = database['Manufacturer/trademark'].astype(str).str.strip().str.lower()
    db_obj = database['Object/part No.'].astype(str).str.strip().str.lower()
    current_obj = ""
    Y_CAP = None
    duplicate = 0

    for idx, cdf_row in cdf.iterrows():

        model_raw = cdf_row.get('Type/model')
        manu_raw = cdf_row.get('Manufacturer/trademark')
        obj_raw = cdf_row.get('Object/part No.')

        model = clean_str(model_raw).lower()
        manu  = clean_str(manu_raw).lower()
        obj = clean_str(obj_raw).lower()


        # 沒提供 model 就很難比對，直接回填原始列
        if not model:
            df = pd.DataFrame([cdf_row], columns=columns)
            output = pd.concat([output, df], ignore_index=True)
            continue

        # 先用 model 模糊比對
        mask = db_model.str.contains(model, case=False, na=False, regex=False)

        # 再視情況加上廠牌條件（只有在 manu 有值時才加)
        if manu:
            mask = mask & db_manu.str.contains(manu, case=False, na=False, regex=False)
        if obj:
            mask = mask & db_obj.str.contains(obj, case=False, na=False, regex=False)

        df = database[mask]

        # 找不到就補原始列
        if df.empty:
            df = pd.DataFrame([cdf_row], columns=columns)

        output = pd.concat([output, df], ignore_index=True)

    output.loc[output["Object/part No."].duplicated(keep="first"), "Object/part No."] = "(Alternate)"

    return output[columns]




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

