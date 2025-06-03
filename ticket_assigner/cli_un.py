# import os
# import glob
# import shutil
# import sqlite3
# import pandas as pd
# from datetime import datetime

# DATA_DIR = "data"
# DB_PATH = os.path.join(DATA_DIR, "ticket_assign.db")

# def find_latest_inm2():
#     downloads_folder = os.path.expandvars(r"%USERPROFILE%\Downloads")
#     pattern = os.path.join(downloads_folder, "INM_202*.xlsx")
#     files = glob.glob(pattern)
#     if not files:
#         raise FileNotFoundError("No INM_*.xlsx files found in Downloads.")
#     latest = max(files, key=os.path.getmtime)
#     target = os.path.join(DATA_DIR, "INM2.xlsx")
#     shutil.copy(latest, target)
#     print(f"📥 Copied latest INM to {target}")
#     return target

# def normalize_xlsx(input_file, output_file):
#     df = pd.read_excel(input_file)
#     df.columns = df.columns.str.strip().str.lower().str.replace(r'\s+', '_', regex=True)
#     df.to_excel(output_file, index=False)
#     return output_file

# def load_unassigned_to_db(xlsx_file):
#     df = pd.read_excel(xlsx_file)
#     conn = sqlite3.connect(DB_PATH)
#     df['created_at'] = datetime.now().isoformat()
#     df.to_sql("unassigned", conn, if_exists='append', index=False)
#     conn.commit()
#     conn.close()
#     print(f"✅ Saved unassigned data to 'unassigned' table")

# def compare_and_export_unassigned():
#     # Load today's INM with blank assignee
#     inm_df = pd.read_excel(os.path.join(DATA_DIR, "INM.normalized.2.xlsx"))
#     inm_df.columns = inm_df.columns.str.strip().str.lower()

#     still_unassigned = inm_df[inm_df['expert_assignee_name'].isna()].copy()

#     # Load previous assignments
#     assigned_df = pd.read_excel(os.path.join(DATA_DIR, "previous_assigned.xlsx"))
#     assigned_df.columns = assigned_df.columns.str.strip().str.lower()

#     # Merge based on ticket ID or description, assuming a column like 'id' or 'summary' exists
#     key_column = 'id' if 'id' in still_unassigned.columns else still_unassigned.columns[0]

#     merged = pd.merge(
#         still_unassigned,
#         assigned_df,
#         how='inner',
#         on=key_column,
#         suffixes=('_new', '_prev')
#     )

#     # Only keep tickets that were assigned earlier
#     result = merged[[col for col in merged.columns if col.endswith('_prev') or col.endswith('_new')]]

#     output_path = os.path.join(DATA_DIR, "TEAM_UnAssigned_Email.xlsx")
#     result.to_excel(output_path, index=False)
#     print(f"📄 Created TEAM_UnAssigned_Email.xlsx with tickets that remain unassigned despite previous assignment.")

# # def main():
# #     latest_inm = find_latest_inm2()
# #     normalized = normalize_xlsx(latest_inm, os.path.join(DATA_DIR, "INM.normalized.2.xlsx"))
# #     load_unassigned_to_db(normalized)
# #     compare_and_export_unassigned()

# # if __name__ == "__main__":
# #     main()

# ticket_assigner/unassigned.py

def cli_un():
    from datetime import datetime
    import os
    import glob
    import shutil
    import pandas as pd
    import sqlite3

    DATA_DIR = os.path.join(os.path.dirname(__file__), "..", "data")
    DB_PATH = os.path.join(DATA_DIR, "ticket_assign.db")

    def find_latest_inm2():
        downloads_folder = os.path.expandvars(r"%USERPROFILE%\Downloads")
        pattern = os.path.join(downloads_folder, "INM_202*.xlsx")
        files = glob.glob(pattern)
        if not files:
            raise FileNotFoundError("No INM_*.xlsx files found in Downloads.")
        latest = max(files, key=os.path.getmtime)
        target = os.path.join(DATA_DIR, "INM2.xlsx")
        shutil.copy(latest, target)
        print(f"📥 Copied latest INM to {target}")
        return target

    def normalize_xlsx(input_file, output_file):
        df = pd.read_excel(input_file)
        df.columns = df.columns.str.strip().str.lower().str.replace(r'\s+', '_', regex=True)
        df.to_excel(output_file, index=False)
        return output_file

    def load_unassigned_to_db(xlsx_file):
        df = pd.read_excel(xlsx_file)
        conn = sqlite3.connect(DB_PATH)
        df['created_at'] = datetime.now().isoformat()
        df.to_sql("unassigned", conn, if_exists='append', index=False)
        conn.commit()
        conn.close()
        print(f"✅ Saved unassigned data to 'unassigned' table")

    def compare_and_export_unassigned():
        inm_df = pd.read_excel(os.path.join(DATA_DIR, "INM.normalized.2.xlsx"))
        inm_df.columns = inm_df.columns.str.strip().str.lower()
        still_unassigned = inm_df[inm_df['expert_assignee_name'].isna()].copy()

        assigned_df = pd.read_excel(os.path.join(DATA_DIR, "previous_assigned.xlsx"))
        assigned_df.columns = assigned_df.columns.str.strip().str.lower()

        key_column = 'id' if 'id' in still_unassigned.columns else still_unassigned.columns[0]

        merged = pd.merge(
            still_unassigned,
            assigned_df,
            how='inner',
            on=key_column,
            suffixes=('_new', '_prev')
        )

        result = merged[[col for col in merged.columns if col.endswith('_prev') or col.endswith('_new')]]
        output_path = os.path.join(DATA_DIR, "TEAM_UnAssigned_Email.xlsx")
        result.to_excel(output_path, index=False)
        print(f"📄 Created TEAM_UnAssigned_Email.xlsx with tickets that remain unassigned despite previous assignment.")

    latest_inm = find_latest_inm2()
    normalized = normalize_xlsx(latest_inm, os.path.join(DATA_DIR, "INM.normalized.2.xlsx"))
    load_unassigned_to_db(normalized)
    compare_and_export_unassigned()

if __name__ == "__main__":
    cli_un()