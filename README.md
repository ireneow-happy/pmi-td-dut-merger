# 📊 AVI PMI + TD Order & DUT# Merge Tool

This Streamlit app is used to integrate AVI PMI inspection files with **TD Order Map** and **DUT# Map** to create enriched result files. It helps engineers analyze probe test data more effectively by adding touchdown order and DUT information.

---

## 📦 Features

- Convert TD / DUT Excel maps to table format
- Merge TD Order and DUT# into AVI PMI Excel files
- Support for uploading multiple AVI PMI files
- Output ZIP containing all merged results

---

## ⚠️ Important Notes

- 📐 **Coordinate system alert**:
  - AVI maps start from the **top-left corner**
  - Prober maps start from the **bottom-left corner**
  - 👉 Please **manually convert all coordinates to top-left zero point** before upload.

- 🧹 **Excel cleanup required**:
  - Each uploaded Excel should contain **only one sheet** to be processed.
  - Please remove any extra sheets manually.

- ⏳ **TD Order Map must be pre-converted** from `st_time` using a different tool.

---

## 🚀 How to Use

1. Run the app:
   ```bash
   streamlit run app.py
   ```

2. Upload the following:
   - TD Order Map `.xlsx`
   - DUT# Map `.xlsx`
   - One or more AVI PMI `.xlsx` files

3. Click the **"🚀 合併並下載結果"** button.

4. Download the merged Excel results as a ZIP file.

---

## 📁 Output

Each input PMI file will generate a new file with additional columns:
- `TD Order`
- `DUT#`

The result will be downloadable as a `.zip` file containing all updated Excel files.

---

## 📦 Requirements

```txt
streamlit
pandas
openpyxl
xlsxwriter
```

---

## 👩‍💻 Developed by

Made by **Irene** for internal engineering use.  
© 2025
