import win32com.client as win32
import os
from bs4 import BeautifulSoup, Tag
import sys
import re # Import regex for cleaning cell text if needed


original_html_file_path = "data\\TEAM_Assigned_Email.html"
html_file_path = "data\\TEAM_Assigned_Email_Formatted.html" 

# Recipient email address
recipient_email = "sumit.das@linde.com"

# Email subject
email_subject = "Daily Incident Ticket"


## Style CSS Start

try:
    # --- Step 1: Re-extract the original table to start fresh ---
    try:
        with open(original_html_file_path, 'r', encoding='utf-8') as f:
            original_html_content = f.read()
        original_soup = BeautifulSoup(original_html_content, 'html.parser')
        original_table = original_soup.find('table')
        if not original_table:
            print("Error: No table found in the original input file.")
            sys.exit(1)
    except FileNotFoundError:
        print(f"Error: Original input file not found at {original_html_file_path}")
        sys.exit(1)

    # --- Step 2: Apply new styles, add header text, and add hyperlinks ---
    soup = BeautifulSoup("", 'html.parser') # Create a new empty soup

    # Create the header paragraph
    header_text = "Daily Incident Assignment"
    header_style = "font-family: Tahoma, sans-serif; font-size: 12pt; text-align: left; margin-bottom: 10px;"
    header_p = soup.new_tag("p", style=header_style)
    header_p.string = header_text

    # Get the table from the original extraction
    table = original_table

    # Define styles with updated font
    table_style = "border: 1px solid black; border-collapse: collapse; width: 75%;"
    font_style = "font-family: Tahoma, sans-serif; font-size: 9pt;"
    th_style = f"border: 1px solid black; padding: 5px; text-align: left; background-color: #a7a8ac; color: black; {font_style}"
    td_style = f"border: 1px solid black; padding: 5px; text-align: left; {font_style}"

    # Apply style to the table tag itself
    table['style'] = table_style

    # Find the index of the 'Id' column
    id_column_index = -1
    headers = table.find_all('th')
    for i, th in enumerate(headers):
        if th.get_text(strip=True).lower() == 'id': # Case-insensitive check
            id_column_index = i
            break

    # Apply style to all th cells
    for cell in headers:
        cell['style'] = th_style

    # Apply style to all td cells and add hyperlinks
    rows = table.find_all('tr')
    for row_index, row in enumerate(rows):
        if row_index == 0 and headers: # Skip header row if headers were found
            continue
        cells = row.find_all('td')
        for i, cell in enumerate(cells):
            cell['style'] = td_style # Apply base style first
            if i == id_column_index:
                # Found the cell in the 'Id' column
                cell_text = cell.get_text(strip=True)
                if cell_text: # Only create link if there's text
                    # Clean text if necessary (e.g., remove extra spaces)
                    cleaned_text = re.sub(r'\s+', '', cell_text).strip()
                    if cleaned_text:
                        link_url = f"https://smax.linde.com/saw/Incident/{cleaned_text}/general"
                        # Clear existing content (important!)
                        cell.clear()
                        # Create new <a> tag
                        link_tag = soup.new_tag('a', href=link_url)
                        link_tag.string = cell_text # Use original text for display
                        # Append the link tag to the cell
                        cell.append(link_tag)
                        # Re-apply style to the cell itself if needed, though inherited style might suffice
                        cell['style'] = td_style

    # --- Step 3: Create the final HTML structure ---
    container_div = soup.new_tag("div")
    container_div.append(header_p)
    container_div.append(table) # Append the modified table

    final_html = str(container_div)

    # Overwrite the extracted_table.html file with the new styled HTML + header + links
    with open(html_file_path, 'w', encoding='utf-8') as out_f:
        out_f.write(final_html)
    print(f"Custom styles, header, font, and Id hyperlinks added to {html_file_path}")

except Exception as e:
    print(f"An error occurred while adding styles/links: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)


## Style CSS End


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