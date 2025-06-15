import pandas as pd

# Load Pre and Post Excel files
pre = pd.read_excel("Pre.xlsx", sheet_name=None)
post = pd.read_excel("Post.xlsx", sheet_name=None)

def compare_dfs(new_df, old_df, sheet_name, id_col='nrBtsID', cellgrp_col='nrCellGrpId'):
    diffs = []
    # اكتشاف الأعمدة معرفات الحالة
    cols_lower = {col.lower(): col for col in new_df.columns}
    bts_col = cols_lower.get(id_col.lower())
    cellgrp = cols_lower.get(cellgrp_col.lower())
    if not bts_col or not cellgrp:
        raise ValueError(f"Required identifier columns '{id_col}' or '{cellgrp_col}' not found.")
    # حد الصفوف والأعمدة للتحقق
    min_rows = min(len(new_df), len(old_df))
    min_cols = min(new_df.shape[1], old_df.shape[1])
    columns = new_df.columns[:min_cols]
    for row in range(min_rows):
        for col in range(min_cols):
            val_new = new_df.iat[row, col]
            val_old = old_df.iat[row, col]
            if pd.isna(val_new) and pd.isna(val_old):
                continue
            if val_new != val_old:
                diffs.append({
                    "Sheet": sheet_name,
                    "Row": row + 2,
                    "nrBtsID": new_df.iat[row, new_df.columns.get_loc(bts_col)],
                    "nrCellGrpId": new_df.iat[row, new_df.columns.get_loc(cellgrp)],
                    "Column": columns[col],
                    "Old Value": val_old,
                    "New Value": val_new
                })
    return diffs

all_diffs = []
for sheet, new_df in post.items():
    old_df = pre.get(sheet)
    if old_df is not None:
        all_diffs.extend(compare_dfs(new_df, old_df, sheet))

# حفظ التقرير
df_diffs = pd.DataFrame(all_diffs)
df_diffs.to_excel("Differences_Report.xlsx", index=False)

print("تمت المقارنة وحفظ التقرير في Differences_Report.xlsx")