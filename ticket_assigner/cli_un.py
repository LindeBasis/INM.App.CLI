import os
import glob
import shutil
import sqlite3
import pandas as pd
from datetime import datetime

DATA_DIR = "data"
# DATA_DIR = os.path.join(os.path.dirname(__file__), "..", "data")
DB_PATH = os.path.join(DATA_DIR, "ticket_assign.db")


def find_latest_inm2():
    downloads_folder = os.path.expandvars(r"%USERPROFILE%\\Downloads")
    pattern = os.path.join(downloads_folder, "Incident_202*.xlsx")
    files = glob.glob(pattern)
    if not files:
        raise FileNotFoundError("No Incident_*.xlsx files found in Downloads.")
    latest = max(files, key=os.path.getmtime)
    target = os.path.join(DATA_DIR, "INM2.xlsx")
    shutil.copy(latest, target)
    print(f"📥 Copied latest INM to {target}")
    return target

def normalize_xlsx(input_file, output_file):
    df = pd.read_excel(input_file)
    df.columns = df.columns.str.strip().str.lower().str.replace(r'\s+', '_', regex=True)
    df.to_excel(output_file, index=False)
    print(f"✅ Normalized file written to {output_file}")
    return output_file

def compare_unassigned_to_previous():
    # Load today's normalized INM2
    inm_df = pd.read_excel(os.path.join(DATA_DIR, "INM.normalized.2.xlsx"))
    inm_df.columns = inm_df.columns.str.strip().str.lower()

    # Only rows where expert_assignee_name is blank
    unassigned_today = inm_df[inm_df['expert_assignee_name'].isna()].copy()

    # Load morning assigned file
    assigned_df = pd.read_excel(os.path.join(DATA_DIR, "TEAM_Assigned_Email.xlsx"))
    assigned_df.columns = assigned_df.columns.str.strip().str.lower()

    # Assume 'id' column uniquely identifies tickets
    key_column = 'id' if 'id' in unassigned_today.columns else unassigned_today.columns[0]

    # Filter assigned_df for only those still unassigned now
    result_df = assigned_df[assigned_df[key_column].isin(unassigned_today[key_column])]

    if result_df.empty:
        print("✅ No lingering unassigned tickets found.")
    else:
        result_path = os.path.join(DATA_DIR, "TEAM_UnAssigned_Email.xlsx")
        result_df.to_excel(result_path, index=False)
        print(f"📤 Created TEAM_UnAssigned_Email.xlsx with {len(result_df)} tickets that remain unassigned.")

def load_unassigned_to_db(xlsx_file):
    xlsx_file_unassigned = os.path.join(DATA_DIR, xlsx_file)
    df = pd.read_excel(xlsx_file_unassigned)
    conn = sqlite3.connect(DB_PATH)
    df['created_at'] = datetime.now().isoformat()
    df.to_sql("unassigned", conn, if_exists='append', index=False)
    conn.commit()
    conn.close()
    print(f"✅ Saved unassigned data to 'unassigned' table")


def generate_html_from_excel(excel_file='TEAM_UnAssigned_Email.xlsx', output_file='TEAM_UnAssigned.html'):
    df = pd.read_excel(os.path.join(DATA_DIR, excel_file), engine='openpyxl')
    html = df.to_html(index=False, border=1)

    with open(os.path.join(DATA_DIR, output_file), 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"✅ HTML report saved to {output_file}")

def cli_un():
    latest_inm = find_latest_inm2()
    normalized = normalize_xlsx(latest_inm, os.path.join(DATA_DIR, "INM.normalized.2.xlsx"))
    compare_unassigned_to_previous()
    load_unassigned_to_db("TEAM_UnAssigned_Email.xlsx")
    generate_html_from_excel()

if __name__ == "__main__":
    cli_un()