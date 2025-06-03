import os
import sqlite3
import pandas as pd
import argparse
from datetime import datetime
import glob
import shutil
import win32com.client as win32
import os

DATA_DIR = "data"
DB_PATH = os.path.join(DATA_DIR, "ticket_assign.db")
html_file_path = "TEAM_Assigned_Email.html" 
recipient_email = "sumit.das@linde.com"
email_subject = "Daily Incident Ticket"


def find_latest_inm_file():
    downloads_folder = os.path.expandvars(r"%USERPROFILE%\Downloads")
    pattern = os.path.join(downloads_folder, "Incident_202*.xlsx")
    matching_files = glob.glob(pattern)

    if not matching_files:
        raise FileNotFoundError("No Incident_202*.xlsx files found in Downloads folder.")

    latest_file = max(matching_files, key=os.path.getmtime)
    print(f"📥 Found latest INM file: {latest_file}")

    target_path = os.path.join(DATA_DIR, "INM.xlsx")
    shutil.copy(latest_file, target_path)
    print(f"✅ Copied latest INM file to: {target_path}")
    return target_path

def normalize_inm_excel(input_file, output_file='INM.normalized.xlsx'):
    input_path = os.path.join(DATA_DIR, input_file)
    output_path = os.path.join(DATA_DIR, output_file)

    df = pd.read_excel(input_path, engine='openpyxl')
    df.columns = (
        df.columns.str.strip()
                  .str.lower()
                  .str.replace(r'\s+', '_', regex=True)
    )
    df.to_excel(output_path, index=False)
    print(f"✅ Normalized Excel written to {output_path}")

def load_inm_excel_to_db(excel_file):
    df = pd.read_excel(os.path.join(DATA_DIR, excel_file), engine='openpyxl')
    df.columns = df.columns.str.strip().str.lower().str.replace(r'\s+', '_', regex=True)

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS tickets (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            created_at TEXT,
            ticket_data TEXT
        )
    """)
    for _, row in df.iterrows():
        cur.execute("INSERT INTO tickets (created_at, ticket_data) VALUES (?, ?)",
                    (datetime.now().isoformat(), row.to_json()))
    conn.commit()
    conn.close()
    print(f"✅ Loaded {len(df)} rows into DB")

def assign_tickets(inm_file, avail_file, team_file, output_file):
    inm_df = pd.read_excel(os.path.join(DATA_DIR, inm_file), engine='openpyxl')
    inm_df.columns = inm_df.columns.str.strip().str.lower().str.replace(r'\s+', '_', regex=True)

    avail_df = pd.read_excel(os.path.join(DATA_DIR, avail_file), engine='openpyxl')
    avail_df.columns = avail_df.columns.str.strip().str.lower().str.replace(r'\s+', '_', regex=True)

    team_df = pd.read_excel(os.path.join(DATA_DIR, team_file), engine='openpyxl')
    team_df.columns = team_df.columns.str.strip().str.lower().str.replace(r'\s+', '_', regex=True)

    if 'status' not in avail_df.columns or 'name' not in avail_df.columns:
        raise ValueError("TEAM_AVAIL.xlsx must contain 'Name' and 'Status' columns")

    available = avail_df[~avail_df['status'].str.lower().eq('leave')]['name'].tolist()
    if not available:
        raise Exception("No team members available for assignment!")

    assigned = []
    member_index = 0
    for _, ticket in inm_df.iterrows():
        member = available[member_index % len(available)]
        member_index += 1
        assigned.append({
            **ticket.to_dict(),
            'assigned_to': member
        })

    assigned_df = pd.DataFrame(assigned)

    # Save TEAM_Assigned.xlsx
    assigned_path = os.path.join(DATA_DIR, output_file)
    assigned_df.to_excel(assigned_path, index=False)
    print(f"✅ Tickets assigned and saved to {output_file}")

    # Keep only unassigned tickets
    unassigned_df = assigned_df[assigned_df['expert_assignee_name'].isna()]

    # Drop email-excluded columns
    columns_to_drop = ['service_display_label', 'expert_assignee_name', 'creation_time']
    unassigned_df = unassigned_df.drop(columns=[col for col in columns_to_drop if col in unassigned_df.columns])

    # Save to TEAM_Assigned_Email.xlsx
    email_path = os.path.join(DATA_DIR, "TEAM_Assigned_Email.xlsx")
    unassigned_df.to_excel(email_path, index=False)
    print(f"📧 Email-ready Excel saved to TEAM_Assigned_Email.xlsx")

    # Generate HTML
    html_path = os.path.join(DATA_DIR, "TEAM_Assigned_Email.html")
    html_content = unassigned_df.to_html(index=False, border=0)
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html_content)
    print(f"📝 HTML email content saved to {html_path}")

    # ✅ Save TEAM_Assigned_Email.xlsx to 'assigned' table in DB
    conn = sqlite3.connect(DB_PATH)
    unassigned_df['created_at'] = datetime.now().isoformat()
    unassigned_df.to_sql("assigned", conn, if_exists='append', index=False)
    conn.commit()
    conn.close()
    print("📦 Saved email-ready data to 'assigned' table in ticket_assign.db")


def generate_html_from_excel(excel_file='TEAM_Assigned_Email.xlsx', output_file='TEAM_Assigned.html'):
    df = pd.read_excel(os.path.join(DATA_DIR, excel_file), engine='openpyxl')
    html = df.to_html(index=False, border=1)

    with open(os.path.join(DATA_DIR, output_file), 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"✅ HTML report saved to {output_file}")

def create_csv_from_db_previousAssigned():
    conn = sqlite3.connect(DB_PATH)
    query = """
    SELECT * FROM assigned
    ORDER BY created_at DESC
    """
    df = pd.read_sql_query(query, conn)
    conn.close()

    if df.empty:
        print("⚠️ No data found in assigned table.")
    else:
        path = os.path.join(DATA_DIR, "previous_assigned.xlsx")
        df.to_excel(path, index=False)
        print(f"✅ previous_assigned.xlsx created with latest entries from 'assigned' table.")


def cli():
    parser = argparse.ArgumentParser(description="Ticket Assigner CLI")
    subparsers = parser.add_subparsers(dest="command")

    # latest_inm_path=find_latest_inm_file()
    # print(latest_inm_path)

    norm_parser = subparsers.add_parser("fetch-latest")
    norm_parser.add_argument('--input', default='INM.xlsx') 
    norm_parser.add_argument('--output', default='INM.normalized.xlsx')

    norm_parser = subparsers.add_parser("normalize-inm")
    norm_parser.add_argument('--input', default='INM.xlsx')
    norm_parser.add_argument('--output', default='INM.normalized.xlsx')

    load_parser = subparsers.add_parser("load-inm")
    load_parser.add_argument('--inm', default='INM.normalized.xlsx')

    assign_parser = subparsers.add_parser("assign")
    assign_parser.add_argument('--inm', default='INM.normalized.xlsx')
    assign_parser.add_argument('--avail', default='TEAM_AVAIL.xlsx')
    assign_parser.add_argument('--team', default='TEAM_MASTER.xlsx')
    assign_parser.add_argument('--output', default='TEAM_Assigned.xlsx')

    html_parser = subparsers.add_parser("html")
    html_parser.add_argument('--input', default='TEAM_Assigned_Email.xlsx')
    html_parser.add_argument('--output', default='TEAM_Assigned.html')

    args = parser.parse_args()


    if args.command == "fetch-latest":
        find_latest_inm_file()
    elif args.command == "normalize-inm":
        normalize_inm_excel(args.input, args.output)
    elif args.command == "load-inm":
        load_inm_excel_to_db(args.inm)
    elif args.command == "assign":
        assign_tickets(args.inm, args.avail, args.team, args.output)
    elif args.command == "html":
        generate_html_from_excel(args.input, args.output)
    else:
        parser.print_help()

if __name__ == "__main__":
    cli()
