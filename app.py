import streamlit as st
import pandas as pd
import zipfile
import io

st.set_page_config(page_title="AVI PMI檔案與TD order及Dut# 資料整合", layout="centered")
st.title("📊 AVI PMI檔案與TD order及Dut# 資料整合")

st.markdown("""
### 📘 功能說明：
1. 將 TD / DUT Map 轉為表格格式
2. 將 TD Order 和 DUT# 合併至AVI PMI 資料檔

⚠️ **程式仍有不完美的地方，請注意：**
- AVI 與 prober 所產生 map 座標零點不同，AVI 在左上角，Prober 在左下角。
  👉 所以請在程式運行前，**手動將所有 Excel 中座標零點改成左上角**。
- 每個上傳的 Excel 檔僅保留需要處理的 sheet，**其餘請移除**。
  👉 程式每次僅處理一個 sheet。
- **TD Order Map 必須先由 st_time 檔轉換**，請使用另一支程式預先處理。
""")
st.markdown("---")  # 區隔線，區分功能說明與檔案上傳區

# === 1. Upload TD/DUT Map ===
st.markdown("### 🔷 TD Order Map")
td_file = st.file_uploader("", type="xlsx", key="td")
st.markdown("<div style='margin-top: -25px'></div>", unsafe_allow_html=True)

st.markdown("### 🔷 DUT# Map")
dut_file = st.file_uploader("", type="xlsx", key="dut")
st.markdown("<div style='margin-top: -25px'></div>", unsafe_allow_html=True)

# === 2. Upload PMI files (multiple) ===
st.markdown("### 🔷 上傳一片或多片AVI PMI Excel 檔案")
pmi_files = st.file_uploader("", type="xlsx", accept_multiple_files=True, key="pmi")
st.markdown("<div style='margin-top: -25px'></div>", unsafe_allow_html=True)

# === Process Button ===
if st.button("🚀 合併並下載結果"):
    if not td_file or not dut_file or not pmi_files:
        st.error("請先上傳所有檔案後再執行。")
    else:
        with st.spinner("處理中，請稍候..."):
            td_df = pd.read_excel(td_file, index_col=0)
            dut_df = pd.read_excel(dut_file, index_col=0)

            td_map = td_df.stack().reset_index()
            td_map.columns = ['Row', 'Col', 'TD Order']
            td_map['Col'] = td_map['Col'].astype(int)

            dut_map = dut_df.stack().reset_index()
            dut_map.columns = ['Row', 'Col', 'DUT#']
            dut_map['Col'] = dut_map['Col'].astype(int)

            zip_buffer = io.BytesIO()
            total_files = len(pmi_files)
            progress_bar = st.progress(0, text="開始合併資料...")

            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
                for idx, uploaded_file in enumerate(pmi_files):
                    st.write(f"🔄 正在處理: {uploaded_file.name}")
                    pmi_df = pd.read_excel(uploaded_file)
                    merged = pmi_df.merge(td_map, on=["Row", "Col"], how="left")
                    merged = merged.merge(dut_map, on=["Row", "Col"], how="left")

                    cols = [c for c in merged.columns if c not in ["TD Order", "DUT#"]] + ["TD Order", "DUT#"]
                    merged = merged[cols]

                    result_bytes = io.BytesIO()
                    merged.to_excel(result_bytes, index=False)
                    result_bytes.seek(0)

                    output_filename = uploaded_file.name.replace(".xlsx", "_with_TD_DUT.xlsx")
                    zipf.writestr(output_filename, result_bytes.read())

                    progress_bar.progress((idx + 1) / total_files, text=f"已完成 {idx + 1} / {total_files} 片")

            zip_buffer.seek(0)
            st.success("✅ 全部處理完成！")
            st.download_button(
                label="📦 下載合併後結果 ZIP 檔",
                data=zip_buffer,
                file_name="PMI_Merge_Results.zip",
                mime="application/zip"
            )