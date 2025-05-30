# ticket_assigner/cli.py

def cli():
    import argparse
    import pandas as pd
    import sqlite3
    from datetime import datetime
    import itertools

    DB_PATH = 'ticket_assign.db'

    def load_inm_to_sqlite(inm_path):
        df_inm = pd.read_csv(inm_path)
        df_inm['created_at'] = datetime.now().isoformat()
        with sqlite3.connect(DB_PATH) as conn:
            df_inm.to_sql('inm_tickets', conn, if_exists='append', index=False)
        print("✅ Loaded INM.csv into SQLite.")

    def assign_tickets(inm_path, avail_path, team_master_path, output_path):
        inm_df = pd.read_csv(inm_path)
        avail_df = pd.read_csv(avail_path)
        master_df = pd.read_csv(team_master_path)
        available = avail_df[~avail_df['Status'].str.lower().eq('leave')]
        if available.empty:
            raise Exception("❌ No team members available.")
        team_ids = available['EmployeeID'].tolist()
        rr_cycle = itertools.cycle(team_ids)
        assigned_rows = []
        for _, ticket in inm_df.iterrows():
            assignee_id = next(rr_cycle)
            assignee = master_df[master_df['EmployeeID'] == assignee_id].iloc[0]
            assigned_rows.append({
                **ticket.to_dict(),
                'AssignedTo': assignee['Name'],
                'Email': assignee['Email']
            })
        assigned_df = pd.DataFrame(assigned_rows)
        assigned_df.to_csv(output_path, index=False)
        print(f"✅ Assigned tickets written to {output_path}")
        return assigned_df

    def assigned_to_html(input_csv, output_html):
        df = pd.read_csv(input_csv)
        html_content = df.to_html(index=False, border=1)
        with open(output_html, 'w') as f:
            f.write(f"<html><body><h2>Today's Ticket Assignments</h2>{html_content}</body></html>")
        print(f"✅ HTML email saved to {output_html}")

    parser = argparse.ArgumentParser(description="Daily Ticket Assignment Tool")
    subparsers = parser.add_subparsers(dest="command")

    load_parser = subparsers.add_parser('load-inm')
    load_parser.add_argument('--inm', required=True, help='Path to INM.csv')

    assign_parser = subparsers.add_parser('assign')
    assign_parser.add_argument('--inm', required=True)
    assign_parser.add_argument('--avail', required=True)
    assign_parser.add_argument('--team', required=True)
    assign_parser.add_argument('--output', default='TEAM_Assigned.csv')

    html_parser = subparsers.add_parser('html')
    html_parser.add_argument('--input', default='TEAM_Assigned.csv')
    html_parser.add_argument('--output', default='TEAM_Assigned.html')

    args = parser.parse_args()

    if args.command == 'load-inm':
        load_inm_to_sqlite(args.inm)
    elif args.command == 'assign':
        assign_tickets(args.inm, args.avail, args.team, args.output)
    elif args.command == 'html':
        assigned_to_html(args.input, args.output)
    else:
        parser.print_help()
