import win32com.client as win32
import os

html_file_path = "data\TEAM_Assigned_Email.html" 

# Recipient email address
recipient_email = "sumit.das@linde.com"

# Email subject
email_subject = "Daily Incident Ticket"

# --- Script Logic ---
def send_email_with_html_table(html_path, to_email, subject):
    """Reads an HTML file and sends its content as the body of an Outlook email."""
    if not os.path.exists(html_path):
        print(f"Error: HTML file not found at {html_path}")
        print("Please update the 'html_file_path' variable in the script.")
        return

    try:
        # Read the HTML content from the file
        with open(html_path, 'r', encoding='utf-8') as f:
            html_body_content = f.read()

        # Create an Outlook application instance
        outlook = win32.Dispatch('outlook.application')

        # Create a new mail item
        mail = outlook.CreateItem(0) # 0 represents olMailItem

        # Set email properties
        mail.To = to_email
        mail.Subject = subject
        # mailto_link = f"mailto:{to_email}?subject={quote(subject)}&body={mailto_body}"
        # Set the body as HTML
        mail.HTMLBody = html_body_content

        # Display the email for review before sending
        # To send directly without review, replace .Display() with .Send()
        mail.Display() 
        # mail.Send()

        print("Outlook email created successfully. Please review and send.")

    except FileNotFoundError:
        print(f"Error: Could not find the HTML file at {html_path}")
    except Exception as e:
        print(f"An error occurred: {e}")
        print("Ensure Outlook is running and you have the 'pywin32' library installed (`pip install pywin32`).")

# --- Execution ---
if __name__ == "__main__":
    send_email_with_html_table(html_file_path, recipient_email, email_subject)