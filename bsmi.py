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

def get_bsmi(cdf, database):
    comp_df = pd.read_excel('component_translate.xlsx')
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
                if col == 'Object/part No.':
                    data = comp_translation(matched.iloc[0][col], comp_df)
                else:
                    text = str(matched.iloc[0][col])
                    # 去除 UL 的部分
                    segs = [s.strip() for s in text.split(",")]
                    
                    if col == 'Standard':
                        ul_count = 0
                        for s in segs:
                            if "UL" in s.upper():
                                ul_count += 1

                        if len(segs) > ul_count:
                            ul_del = True
                        else:
                            ul_del = False

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
                    else:
                        ss = []
                        for s in segs:
                            s = s.strip()
                            if all(x not in s.upper() for x in ["UL", "EDITION"]):
                                if ":" in s or "：" in s:
                                    s = s.split(":")[0].strip()
                                ss.append(s)
                    data = ", ".join(ss)

                cdf.at[idx, col] = data
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
    print(df.head())


    # df = pd.read_excel(url, sheet_name='CDF_database_2025.03.27_Chris')
    
    cdf_df = pd.read_excel(cdf_path)
    cdf_df = get_bsmi(cdf_df, df)
    output = BytesIO()
    cdf_df.to_excel(output, index=False)
    output.seek(0)
    return output
