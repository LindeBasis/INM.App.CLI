# ticket_assigner/cli.py

def cli():
    import argparse
    import pandas as pd
    import sqlite3
    from datetime import datetime
    import itertools
    import os

    DB_PATH = 'ticket_assign.db'
    DATA_DIR = 'data'

    def load_inm_to_sqlite(inm_filename):
        inm_path = os.path.join(DATA_DIR, inm_filename)
        df_inm = pd.read_csv(inm_path)
        df_inm['created_at'] = datetime.now().isoformat()
        with sqlite3.connect(DB_PATH) as conn:
            df_inm.to_sql('inm_tickets', conn, if_exists='append', index=False)
        print(f"✅ Loaded {inm_filename} into SQLite.")

    def assign_tickets(inm_file, avail_file, team_file, output_file):
        inm_df = pd.read_csv(os.path.join(DATA_DIR, inm_file))
        avail_df = pd.read_csv(os.path.join(DATA_DIR, avail_file))
        master_df = pd.read_csv(os.path.join(DATA_DIR, team_file))

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
        assigned_df.to_csv(os.path.join(DATA_DIR, output_file), index=False)
        print(f"✅ Tickets written to {output_file}")
        return assigned_df

    def assigned_to_html(input_file, output_html):
        df = pd.read_csv(os.path.join(DATA_DIR, input_file))
        html_content = df.to_html(index=False, border=1)
        with open(os.path.join(DATA_DIR, output_html), 'w') as f:
            f.write(f"<html><body><h2>Today's Ticket Assignments</h2>{html_content}</body></html>")
        print(f"✅ HTML file created: {output_html}")

    parser = argparse.ArgumentParser(description="Ticket Assignment CLI")
    subparsers = parser.add_subparsers(dest="command")

    load_parser = subparsers.add_parser('load-inm')
    load_parser.add_argument('--inm', required=True, help='INM CSV file in data/')

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
