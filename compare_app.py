
import streamlit as st
import pandas as pd
import tempfile

def compare_dataframes(new_df, old_df, key_column, sheet_name):
    changes = []
    added_rows = []
    removed_rows = []

    if key_column not in new_df.columns or key_column not in old_df.columns:
        return changes, added_rows, removed_rows, f"❌ العمود '{key_column}' غير موجود في الشيت {sheet_name}"

    new_df.set_index(key_column, inplace=True)
    old_df.set_index(key_column, inplace=True)

    common_columns = new_df.columns.intersection(old_df.columns)
    common_ids = new_df.index.intersection(old_df.index)

    for idx in common_ids:
        for col in common_columns:
            new_val = new_df.at[idx, col]
            old_val = old_df.at[idx, col]
            if pd.isna(new_val) and pd.isna(old_val):
                continue
            if new_val != old_val:
                changes.append({
                    "Sheet": sheet_name,
                    "ID": idx,
                    "Column": col,
                    "Old Value": old_val,
                    "New Value": new_val
                })

    added_ids = new_df.index.difference(old_df.index)
    for idx in added_ids:
        row = new_df.loc[idx].to_dict()
        row.update({"Sheet": sheet_name, "ID": idx})
        added_rows.append(row)

    removed_ids = old_df.index.difference(new_df.index)
    for idx in removed_ids:
        row = old_df.loc[idx].to_dict()
        row.update({"Sheet": sheet_name, "ID": idx})
        removed_rows.append(row)

    return changes, added_rows, removed_rows, None

st.set_page_config(page_title="مقارنة ملفات Excel/CSV", layout="centered")
st.title("📊 مقارنة ملفات Excel أو CSV")

uploaded_new = st.file_uploader("📤 حمل الملف الجديد (Excel أو CSV)", type=["xlsx", "csv"])
uploaded_old = st.file_uploader("📤 حمل الملف القديم (Excel أو CSV)", type=["xlsx", "csv"])

key_column = st.text_input("🔑 أدخل اسم العمود المفتاح للمقارنة", value="ID")

if uploaded_new and uploaded_old:
    compare_button = st.button("🔍 قارن الملفات")
    if compare_button:
        try:
            def load_file(file):
                if file.name.endswith(".csv"):
                    return {"CSV": pd.read_csv(file)}
                else:
                    return pd.read_excel(file, sheet_name=None)

            new_sheets = load_file(uploaded_new)
            old_sheets = load_file(uploaded_old)

            all_changes, all_added, all_removed = [], [], []
            for sheet_name in new_sheets:
                old_sheet_name = sheet_name if sheet_name in old_sheets else list(old_sheets.keys())[0]
                changes, added, removed, error = compare_dataframes(
                    new_sheets[sheet_name],
                    old_sheets.get(old_sheet_name, pd.DataFrame()),
                    key_column,
                    sheet_name
                )
                if error:
                    st.warning(error)
                    continue
                all_changes.extend(changes)
                all_added.extend(added)
                all_removed.extend(removed)

            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                output_path = tmp.name

            with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
                workbook = writer.book
                yellow_fmt = workbook.add_format({'bg_color': '#FFFF00'})
                green_fmt = workbook.add_format({'bg_color': '#C6EFCE'})
                red_fmt = workbook.add_format({'bg_color': '#FFC7CE'})

                if all_changes:
                    df1 = pd.DataFrame(all_changes)
                    df1.to_excel(writer, index=False, sheet_name='Cell Changes')
                    ws = writer.sheets['Cell Changes']
                    for r in range(1, len(df1) + 1):
                        ws.set_row(r, None, yellow_fmt)

                if all_added:
                    df2 = pd.DataFrame(all_added)
                    df2.to_excel(writer, index=False, sheet_name='Added Rows')
                    ws = writer.sheets['Added Rows']
                    for r in range(1, len(df2) + 1):
                        ws.set_row(r, None, green_fmt)

                if all_removed:
                    df3 = pd.DataFrame(all_removed)
                    df3.to_excel(writer, index=False, sheet_name='Removed Rows')
                    ws = writer.sheets['Removed Rows']
                    for r in range(1, len(df3) + 1):
                        ws.set_row(r, None, red_fmt)

            st.success("✅ تمت مقارنة الملفات بنجاح!")
            with open(output_path, "rb") as f:
                st.download_button("⬇️ تحميل ملف النتائج", f, file_name="Differences_Report.xlsx")

        except Exception as e:
            st.error(f"❌ حدث خطأ أثناء المعالجة: {str(e)}")
