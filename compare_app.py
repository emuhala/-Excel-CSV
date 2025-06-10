
import streamlit as st
import pandas as pd
import tempfile

def compare_sheets_by_position(new_df, old_df, sheet_name):
    differences = []

    min_rows = min(len(new_df), len(old_df))
    min_cols = min(len(new_df.columns), len(old_df.columns))
    columns = new_df.columns[:min_cols]

    for row in range(min_rows):
        for col in range(min_cols):
            new_val = new_df.iloc[row, col]
            old_val = old_df.iloc[row, col]
            if pd.isna(new_val) and pd.isna(old_val):
                continue
            if new_val != old_val:
                differences.append({
                    "Sheet": sheet_name,
                    "Row": row + 2,  # Excel rows (header + 1-indexed)
                    "Column": columns[col],
                    "Old Value": old_val,
                    "New Value": new_val
                })
    return differences

st.set_page_config(page_title="مقارنة Excel و CSV بدون ID", layout="centered")
st.title("📊 مقارنة كل خلية في ملفات Excel أو CSV")

uploaded_new = st.file_uploader("📥 الملف الجديد (Excel أو CSV)", type=["xlsx", "csv"])
uploaded_old = st.file_uploader("📥 الملف القديم (Excel أو CSV)", type=["xlsx", "csv"])

if uploaded_new and uploaded_old:
    if st.button("🔍 قارن الملفات"):
        try:
            def load_file(file):
                if file.name.endswith(".csv"):
                    return {"CSV": pd.read_csv(file)}
                else:
                    return pd.read_excel(file, sheet_name=None)

            new_sheets = load_file(uploaded_new)
            old_sheets = load_file(uploaded_old)

            all_diffs = []

            for sheet_name in new_sheets:
                if sheet_name in old_sheets:
                    new_df = new_sheets[sheet_name]
                    old_df = old_sheets[sheet_name]
                    diffs = compare_sheets_by_position(new_df, old_df, sheet_name)
                    all_diffs.extend(diffs)
                else:
                    st.warning(f"❌ الشيت '{sheet_name}' غير موجود في الملف القديم.")

            if all_diffs:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    output_path = tmp.name

                with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
                    workbook = writer.book
                    yellow_fmt = workbook.add_format({'bg_color': '#FFFF00'})
                    df = pd.DataFrame(all_diffs)
                    df.to_excel(writer, index=False, sheet_name='Differences')
                    ws = writer.sheets['Differences']
                    for r in range(1, len(df) + 1):
                        ws.set_row(r, None, yellow_fmt)

                st.success("✅ تم استخراج الفروقات بين الخلايا")
                with open(output_path, "rb") as f:
                    st.download_button("⬇️ تحميل تقرير الفروقات", f, file_name="Cell_Differences_Report.xlsx")
            else:
                st.info("✅ لا توجد فروقات بين الملفات")

        except Exception as e:
            st.error(f"❌ خطأ أثناء المقارنة: {e}")
