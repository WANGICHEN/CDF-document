import pandas as pd
from docx import Document
from docx.shared import RGBColor
import re
from docx.shared import Pt
import requests
from io import BytesIO


columns = [
    'Object/part No.', 'Manufacturer/trademark', 'Type/model',
    'Technical data', 'Standard', 'Mark(s) of conformity', 'website (UL)', 'VDE/TUV/ENEC/BSMI'
]

def clean_str(x):
    if pd.isna(x):
        return ""
    s = str(x).strip()
    return "" if s.lower() in {"nan", "none"} else s

def comp_translation(comp, comp_df):
    result = comp_df[comp_df['英文名稱'] == comp]
    result = comp_df[comp_df['英文名稱'].str.contains(comp, case=False, na=False)]
    if result.empty:
        return comp
    else:
        return result['中文名稱'].values[0]
    
def count_ul(segs):
    ul_count = 0
    for s in segs:
        if "UL" in s.upper():
            ul_count += 1

    if len(segs) > ul_count:
        ul_del = True
    else:
        ul_del = False

    return ul_del

def del_ul_edition(segs, ul_del = True):
    ss = []
    for s in segs:
        s = s.strip()
        if ul_del:
            if all(x not in s.upper() for x in ["UL", "EDITION"]):
                if ":" in s or "：" in s:
                    s = s.split(":")[0].strip()
                ss.append(s)
        else:
            ss.append(s)
    return ss

def clean_data(df, trans_df):
    # comp_df = pd.read_excel('component_translate.xlsx')
    for idx, cdf_row in df.iterrows():
        for col in ['Object/part No.', 'Technical data', 'Standard', 'Mark(s) of conformity', 'website (UL)', 'VDE/TUV/ENEC/BSMI']:
            if col == 'Object/part No.':
                data = comp_translation(cdf_row[col], trans_df)
            else:
                text = str(cdf_row[col])
                if col not in ['website (UL)', 'VDE/TUV/ENEC/BSMI']:
                    # 去除 UL 的部分
                    segs = [s.strip() for s in text.split(",")]
                    if col == 'Standard':
                        ul_del = count_ul(segs)
    
                        ss = del_ul_edition(segs, ul_del)
                    else:
                        ss = del_ul_edition(segs)
                    data = ", ".join(ss)
                else:
                    data = text

            df.at[idx, col] = data
    return df

def get_bsmi(cdf, database):

    ## translation data
    share_url = "https://z28856673-my.sharepoint.com/:x:/g/personal/itek_project_i-tek_com_tw/EV_kkQot_hZNsD8LFYQLfqoBWvR5p28e8_7yvoQXeVHtkg?e=ucQuNr"

    # 多數 SharePoint 連結加上 download=1 可直下（若仍 403，請看下方 路徑B）
    download_url = share_url + ("&download=1" if "?" in share_url else "?download=1")

    headers = {"User-Agent": "Mozilla/5.0"}
    r = requests.get(download_url, headers=headers, allow_redirects=True, timeout=30)
    r.raise_for_status()  # 403/404 會在這裡丟錯

    trans_df = pd.read_excel(BytesIO(r.content), sheet_name=0)  # 或指定 sheet_name
    ##

    
    output = pd.DataFrame(columns=columns)

    # 預先把 database 要比對的欄位正規化（加速與一致性）
    db_model = database['Type/model'].astype(str).str.strip().str.lower()
    db_manu  = database['Manufacturer/trademark'].astype(str).str.strip().str.lower()

    for idx, cdf_row in cdf.iterrows():

        manu_raw  = cdf_row.get('Manufacturer/trademark')
        model_raw = cdf_row.get('Type/model')

        manu  = clean_str(manu_raw).lower()
        model = clean_str(model_raw).lower()

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

        df = database[mask]

        # 找不到就補原始列
        if df.empty:
            df = pd.DataFrame([cdf_row], columns=columns)

        output = pd.concat([output, clean_data(df, trans_df)], ignore_index=True)

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
    cdf_df = get_bsmi(cdf_df, df)
    output = BytesIO()
    cdf_df.to_excel(output, index=False)
    output.seek(0)
    return output









