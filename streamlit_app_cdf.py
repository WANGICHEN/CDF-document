import streamlit as st
import os
import cdf
import bsmi
import tempfile

st.title("PDF → Word 自動轉換工具")
bsmi_on = st.toggle("BSMI")
# 上傳 PDF
cdf_file = st.file_uploader("請上傳 CDF 檔案", type=["xlsx"])


if cdf_file:
    # 檢查檔案是否為 Excel 檔案
    if not cdf_file.name.endswith(".xlsx"): 
        st.error("請上傳一個有效的 Excel 檔案 (.xlsx)")
    else:
        with tempfile.TemporaryDirectory() as tmpdir:
            # 儲存 PDF
            cdf_path = os.path.join(tmpdir, cdf_file.name)
            # Save the uploaded file to cdf_path
            with open(cdf_path, "wb") as f:
                f.write(cdf_file.getbuffer())
            download_buttons = []
            
            word_output_name = cdf_file.name.replace(".xlsx", f".docx")
            output_path = os.path.join(tmpdir, word_output_name)
            if bsmi_on:
                doc = bsmi.run(cdf_path)
            else:
                doc = cdf.run(cdf_path)
            doc.save(output_path)

            with open(output_path, "rb") as out_file:
                st.download_button(
                    label=f"下載 Word 檔",
                    data=out_file,
                    file_name=word_output_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
