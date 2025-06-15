import pandas as pd

# Load Pre and Post Excel files (ensure they are in the same directory as this script)
pre_files = pd.read_excel("Pre.xlsx", sheet_name=None)
post_files = pd.read_excel("Post.xlsx", sheet_name=None)

def compare_sheets(old_df, new_df, sheet_name, id_col="nrBtsID", grp_col="nrCellGrpId"):
    diffs = []
    cols_map = {col.lower(): col for col in new_df.columns}
    try:
        id_actual = cols_map[id_col.lower()]
        grp_actual = cols_map[grp_col.lower()]
    except KeyError as e:
        raise KeyError(f"Missing identifier column: {e.args[0]}")

    max_rows = min(len(old_df), len(new_df))
    max_cols = min(old_df.shape[1], new_df.shape[1])
    headers = new_df.columns[:max_cols]

    for r in range(max_rows):
        for c in range(max_cols):
            old_val = old_df.iat[r, c]
            new_val = new_df.iat[r, c]
            if pd.isna(old_val) and pd.isna(new_val):
                continue
            if old_val != new_val:
                diffs.append({
                    "Sheet": sheet_name,
                    "Row": r + 2,
                    "nrBtsID": new_df.iat[r, new_df.columns.get_loc(id_actual)],
                    "nrCellGrpId": new_df.iat[r, new_df.columns.get_loc(grp_actual)],
                    "Column": headers[c],
                    "Old Value": old_val,
                    "New Value": new_val
                })
    return diffs

# Main execution
if __name__ == "__main__":
    all_diffs = []
    for sheet_name in pre_files:
        if sheet_name in post_files:
            diffs = compare_sheets(pre_files[sheet_name], post_files[sheet_name], sheet_name)
            all_diffs.extend(diffs)

    if all_diffs:
        df = pd.DataFrame(all_diffs)
        report = "Differences_Report.xlsx"
        df.to_excel(report, index=False)
        print(f"Generated {report} with {len(df)} differences.")
    else:
        print("No differences found.")