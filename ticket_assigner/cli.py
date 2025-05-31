import os
import sqlite3
import pandas as pd
import argparse
from datetime import datetime

DATA_DIR = "data"
DB_PATH = os.path.join(DATA_DIR, "ticket_assign.db")


def normalize_inm_csv(input_file, output_file='INM.normalized.csv'):
    input_path = os.path.join(DATA_DIR, input_file)
    output_path = os.path.join(DATA_DIR, output_file)

    df = pd.read_csv(input_path)
    df.columns = (
        df.columns.str.strip()
                  .str.lower()
                  .str.replace(r'\s+', '_', regex=True)
    )
    df.to_csv(output_path, index=False)
    print(f"✅ Normalized CSV written to {output_path}")


def load_inm_csv_to_db(csv_file):
    df = pd.read_csv(os.path.join(DATA_DIR, csv_file))
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
    inm_df = pd.read_csv(os.path.join(DATA_DIR, inm_file))
    inm_df.columns = inm_df.columns.str.strip().str.lower().str.replace(r'\s+', '_', regex=True)

    avail_df = pd.read_csv(os.path.join(DATA_DIR, avail_file))
    avail_df.columns = avail_df.columns.str.strip().str.lower().str.replace(r'\s+', '_', regex=True)

    team_df = pd.read_csv(os.path.join(DATA_DIR, team_file))
    team_df.columns = team_df.columns.str.strip().str.lower().str.replace(r'\s+', '_', regex=True)

    if 'status' not in avail_df.columns or 'name' not in avail_df.columns:
        raise ValueError("TEAM_AVAIL.csv must contain 'Name' and 'Status' columns")

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

    # Save TEAM_Assigned.csv
    assigned_path = os.path.join(DATA_DIR, output_file)
    assigned_df.to_csv(assigned_path, index=False)
    print(f"✅ Tickets assigned and saved to {output_file}")

    # Create TEAM_Assigned_Email.csv
    # ✅ Step 1: Remove rows with NaN in 'expert_assignee_name'
    cleaned_df = assigned_df.dropna(subset=['expert_assignee_name'])

    # ✅ Step 2: Drop unnecessary columns for the email report
    columns_to_drop = ['service_display_label', 'expert_assignee_name', 'creation_time']
    cleaned_df = cleaned_df.drop(columns=[col for col in columns_to_drop if col in cleaned_df.columns])

    # ✅ Step 3: Save the final cleaned CSV
    email_path = os.path.join(DATA_DIR, "TEAM_Assigned_Email.csv")
    cleaned_df.to_csv(email_path, index=False)
    print(f"📧 Final email-ready CSV saved to TEAM_Assigned_Email.csv")


def generate_html_from_csv(csv_file='TEAM_Assigned_Email.csv', output_file='TEAM_Assigned.html'):
    df = pd.read_csv(os.path.join(DATA_DIR, csv_file))
    html = df.to_html(index=False, border=1)

    with open(os.path.join(DATA_DIR, output_file), 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"✅ HTML report saved to {output_file}")


def cli():
    parser = argparse.ArgumentParser(description="Ticket Assigner CLI")
    subparsers = parser.add_subparsers(dest="command")

    # normalize-inm
    norm_parser = subparsers.add_parser("normalize-inm")
    norm_parser.add_argument('--input', default='INM.csv')
    norm_parser.add_argument('--output', default='INM.normalized.csv')

    # load-inm
    load_parser = subparsers.add_parser("load-inm")
    load_parser.add_argument('--inm', default='INM.normalized.csv')

    # assign
    assign_parser = subparsers.add_parser("assign")
    assign_parser.add_argument('--inm', default='INM.normalized.csv')
    assign_parser.add_argument('--avail', default='TEAM_AVAIL.csv')
    assign_parser.add_argument('--team', default='TEAM_MASTER.csv')
    assign_parser.add_argument('--output', default='TEAM_Assigned.csv')

    # html
    html_parser = subparsers.add_parser("html")
    html_parser.add_argument('--input', default='TEAM_Assigned_Email.csv')
    html_parser.add_argument('--output', default='TEAM_Assigned.html')

    args = parser.parse_args()

    if args.command == "normalize-inm":
        normalize_inm_csv(args.input, args.output)
    elif args.command == "load-inm":
        load_inm_csv_to_db(args.inm)
    elif args.command == "assign":
        assign_tickets(args.inm, args.avail, args.team, args.output)
    elif args.command == "html":
        generate_html_from_csv(args.input, args.output)
    else:
        parser.print_help()


if __name__ == "__main__":
    cli()
