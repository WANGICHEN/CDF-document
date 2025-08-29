import streamlit as st
import os
import cdf
import bsmi
import tempfile
from docx import Document
import transfer_word as transfer
import requests
from io import BytesIO  

st.title("零件用料轉換工具")
search = st.checkbox("零件用料轉換")
to_word = st.checkbox("文件檔傳換")

bsmi_on = st.toggle("BSMI")
# 上傳 PDF
cdf_file = st.file_uploader("請上傳想要轉換的檔案", type=["xlsx"])


if cdf_file:
    # 檢查檔案是否為 Excel 檔案
    if not cdf_file.name.endswith(".xlsx"): 
        st.error("請上傳一個有效的 Excel 檔案 (.xlsx)")
    else:
        with tempfile.TemporaryDirectory() as tmpdir:
            cdf_path = os.path.join(tmpdir, cdf_file.name)
            # Save the uploaded file to cdf_path
            with open(cdf_path, "wb") as f:
                f.write(cdf_file.getbuffer())
            download_buttons = []
            
            word_output_name = cdf_file.name.replace(".xlsx", f".docx")
            output_path = os.path.join(tmpdir, word_output_name)

            if search:
                if bsmi_on:
                    excel_bytes = bsmi.run(cdf_path)
                else:
                    excel_bytes = cdf.run(cdf_path)
                st.download_button(
                    label="下載 Excel 檔",
                    data=excel_bytes,
                    file_name=cdf_file.name.replace(".xlsx", "_output.xlsx"),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                
            elif to_word:
                if bsmi_on:
                    f = "BSMI_CDF_template.docx"
                else:
                    f = "UL_CDF_template.docx"
                doc = Document(f)
                output = transfer.WriteInDataSheet(doc, cdf_path)
                buffer = BytesIO()
                output.save(buffer)
                buffer.seek(0)

                st.download_button(
                    label=f"下載 Word 檔",
                    data=buffer,
                    file_name=word_output_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


