# requirements: public
from fuzzywuzzy import fuzz
import re
import win32com.client
import tkinter as tk
import os
import string

# requirements: custom functions
from utils import read_params_from_txt_file
from utils import read_excel_or_pickle

# Function to search for an email in the Excel file
def search_email(subject, sender, recipients):
    filtered_df = df.loc[(df['Subject'] == subject) & (df['Sender'] == sender) & (df['Recipients'] == recipients)]
    return filtered_df

# Function to search for an email in the Excel file, only for subject
def search_email_subject(subject):
    filtered_df = df.loc[(df['Subject'] == subject)]
    return filtered_df

# Function to find the top 3 most likely subjects and corresponding folders
def find_top_matches(subject, sender, recipients):
    # Calculate the similarity score for the email subjects using fuzz.token_set_ratio
    df['Subject_Score'] = df['Subject'].apply(lambda x: fuzz.token_set_ratio(subject, x))

    # Calculate the similarity score for the sender and recipients, considering the role switch
    # We calculate the similarity score for sender and recipient pairs and take the maximum score
    df['Send_Recv_Score'] = df.apply(
        lambda x: max(fuzz.token_set_ratio(sender, x['Sender']) + fuzz.token_set_ratio(recipients, x['Recipients']),
                      fuzz.token_set_ratio(sender, x['Recipients']) + fuzz.token_set_ratio(recipients, x['Sender'])),
        axis=1)

    # Check for a match between the email subject and the folder name
    df['Folder_Name_Score'] = df['Path'].apply(lambda x: fuzz.token_set_ratio(subject, os.path.basename(x)))

    # Calculate the total score by giving equal weight to the subject similarity, the sender/receiver similarity, and folder name similarity
    df['Total_Score'] = df['Subject_Score'] * 0.4 + df['Send_Recv_Score'] * 0.4 + df['Folder_Name_Score'] * 0.2

    # Get the top matches based on the total score
    top_matches = df.nlargest(50, 'Total_Score')

    # Remove duplicates folder paths and keep the top 3 unique items
    top_matches = top_matches.drop_duplicates(subset='Path', keep='first').head(3)

    return top_matches

# Function to open the folder in Windows Explorer
def open_folder(folder_path):
    os.startfile(folder_path)

# Function to show a popup confirmation before archiving
def show_yes_no_popup(prompt):
    def on_yes():
        nonlocal user_choice
        user_choice.set("yes")
        window.destroy()

    def on_no():
        nonlocal user_choice
        user_choice.set("no")
        window.destroy()

    def on_open_folder():
        nonlocal user_choice
        user_choice.set("open_folder")
        window.destroy()

    window = tk.Tk()
    window.title("Confirm if you want to proceed")

    # Set the window style to look like a native Windows dialog
    window.attributes('-toolwindow', True)
    window.lift()
    window.focus_force()
    window.resizable(False, False)
    window.config(padx=10, pady=10)

    # Center the window
    window_width = 900
    window_height = 100
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width / 2) - (window_width / 2)
    y = (screen_height / 2) - (window_height / 2)
    window.geometry(f"{window_width}x{window_height}+{int(x)}+{int(y)}")

    tk.Label(window, text=prompt).pack(pady=10)

    # Create a new frame for the buttons
    button_frame = tk.Frame(window)
    button_frame.pack(pady=10)

    yes_button = tk.Button(button_frame, text="Yes, archive", command=on_yes)
    yes_button.pack(side=tk.LEFT, padx=(0, 5))

    no_button = tk.Button(button_frame, text="No", command=on_no)
    no_button.pack(side=tk.LEFT, padx=(5, 5))

    open_folder_button = tk.Button(button_frame, text="Open Folder", command=on_open_folder)
    open_folder_button.pack(side=tk.LEFT, padx=(5, 0))

    user_choice = tk.StringVar()
    window.mainloop()
    return user_choice.get()

# Function to get the selected email from Outlook
def get_selected_email():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    explorer = outlook.Application.ActiveExplorer()
    selection = explorer.Selection

    if len(selection) == 0:
        print("No email is selected.")
        return None

    return selection.Item(1)

# Function to create a popup window with an input entry
def show_input_popup(prompt, options):
    def on_submit():
        nonlocal user_input
        user_input.set(input_entry.get())
        window.destroy()

    window = tk.Tk()
    window.title("Choose an option")

    # Set the window style to look like a native Windows dialog
    window.attributes('-toolwindow', True)
    window.lift()
    window.focus_force()
    window.resizable(False, False)
    window.config(padx=10, pady=10)

    # Center the window
    window_width = 900
    window_height = 200
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width / 2) - (window_width / 2)
    y = (screen_height / 2) - (window_height / 2)
    window.geometry(f"{window_width}x{window_height}+{int(x)}+{int(y)}")

    tk.Label(window, text=prompt, anchor=tk.W).pack(anchor=tk.W, pady=(0, 10))

    # Display the folder options
    for option in options:
        folder_path = option.split("(Folder: ")[1].rstrip(")")
        tk.Label(window, text=folder_path, anchor=tk.W).pack(anchor=tk.W)

    input_entry = tk.Entry(window)
    input_entry.pack(pady=(15, 10))

    # Set focus to the input entry
    input_entry.focus_set()

    submit_button = tk.Button(window, text="Submit", command=on_submit)
    submit_button.pack(pady=(10, 15))

    user_input = tk.StringVar()
    window.mainloop()
    return user_input.get()

# Function to get the next correlative number in the folder
def get_next_correlative_number(folder_path):
    files = os.listdir(folder_path)
    correlative_numbers = [int(re.findall(r'\d+', f)[0]) for f in files if re.findall(r'\d+', f)]

    if correlative_numbers:
        next_number = max(correlative_numbers) + 1
    else:
        next_number = 1

    return f"{next_number:03d}"

# Function to sanitize the email subject
def sanitize_subject(subject):
    valid_chars = "-_.() %s%s" % (string.ascii_letters, string.digits)
    sanitized_subject = ''.join(c for c in subject if c in valid_chars)
    return sanitized_subject.strip()

# Function to save email attachments
def save_attachments(email, folder_path, correlative_number):
    for attachment in email.Attachments:
        try:
            content_id = attachment.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E")
        except Exception:
            content_id = None

        # Check if the attachment has a ContentId, which is common for embedded images
        if content_id:
            continue

        file_path = os.path.join(folder_path, f"{correlative_number} - {attachment.FileName}")
        attachment.SaveAsFile(file_path)

# Function to save the email as a .msg file
def save_email_as_msg(email, folder_path, correlative_number):
    sanitized_subject = sanitize_subject(email.Subject)
    file_name = f"{correlative_number} - {sanitized_subject}.msg"
    file_path = os.path.join(folder_path, file_name)
    email.SaveAs(file_path)

# Function to archive the email
def archive_email(email):
    # Get the Gmail account
    outlook = email.Application.GetNamespace("MAPI")
    accounts = outlook.Accounts
    gmail_account = None

    for account in accounts:
        if "gmail" in account.SmtpAddress.lower():
            gmail_account = account
            break

    if not gmail_account:
        print("No Gmail account found.")
        return

    # Get the "All Mail" folder, which is the archive folder for Gmail
    all_mail_folder = gmail_account.DeliveryStore.GetRootFolder().Folders["[Gmail]"].Folders["Archivo"]

    if not all_mail_folder:
        print("No Gmail 'All Mail' folder found.")
        return

    # Move the email to the Gmail "All Mail" folder
    email.Move(all_mail_folder)

# Main execution

# Load the parameters from the text file
params_file_path = r"C:\Mis Datos en Local\temporal\python\email-automation-archive-params.txt"
params = read_params_from_txt_file(params_file_path)

# Load the existing Excel file as a DataFrame
excel_path = params['excel_path']
pickle_path = params['pickle_path']

df = read_excel_or_pickle(excel_path,pickle_path)

# Evaluate the email and decide
email = get_selected_email()

if email:

    # Remove "re" or "rv" prefix and any leading/trailing whitespace or ".msg" suffix from email subject
    subject = re.sub(r"^\s*(re|rv|fw[d]*)\s*:?\s*|\s*\.msg$", "", email.Subject, flags=re.IGNORECASE)

    sender = email.SenderEmailAddress

    # try to get the PrimarySmtpAddress from the Exchange user only if it exists, otherwise, it will use the original address
    recipients = ';'.join([r.Address if r.Type != 0 else (r.AddressEntry.GetExchangeUser().PrimarySmtpAddress if r.AddressEntry.GetExchangeUser() else r.Address) for r in email.Recipients])

    print('subject = ' + subject + ' sender = ' + sender + ' - recipients = ' + recipients)

    match = search_email(subject, sender, recipients)
    if not match.empty:
        print(f"Perfect match by subject, sender, recipient found. The folder path is: {folder_path}")

        # Show a Yes/No popup asking if the user wants to proceed with the perfect match
        prompt = f"Perfect match by subject, sender, recipient found. The folder path is:\n\n {match.iloc[0]['Path']}\n\n Do you want to proceed?"
        proceed = show_yes_no_popup(prompt)

        if proceed == "yes":
            folder_path = match.iloc[0]['Path']
            print(f"Proceeding with perfect match. The folder path is: {folder_path}")
        elif proceed == "open_folder":
            folder_path = match.iloc[0]['Path']
            open_folder(folder_path)
            print("Folder opened - archive manually")
            exit()
        else:
            folder_path = None
            print("User chose not to proceed with the perfect match.")
    else:
        match = search_email_subject(subject)
        if not match.empty:
            folder_path = match.iloc[0]['Path']
            print(f"Match by subject found. The folder path is: {folder_path}")

            # Show a Yes/No popup asking if the user wants to proceed with the perfect match
            prompt = f"Subject match. The folder path is:\n\n {match.iloc[0]['Path']}\n\n Do you want to proceed?"
            proceed = show_yes_no_popup(prompt)

            if proceed == "yes":
                folder_path = match.iloc[0]['Path']
                print(f"Proceeding with perfect match. The folder path is:\n {folder_path}")
            elif proceed == "open_folder":
                folder_path = match.iloc[0]['Path']
                open_folder(folder_path)
                print("Folder opened - archive manually")
                exit()
            else:
                folder_path = None
                print("User chose not to proceed with the subject match, checking top matches.")

                top_matches = find_top_matches(subject, sender, recipients)

                # print the three top options
                print("Top 3 matches:")
                for idx, row in top_matches.iterrows():
                    print(f"{idx + 1}: {row['Subject']} (Folder: {row['Path']})")

                # Get the user's choice from the popup
                options = [f"{idx + 1}: {row['Subject']} (Folder: {row['Path']})" for idx, row in
                           top_matches.iterrows()]
                prompt = "Enter the number of the chosen option (1/2/3) or 'o' to open the first folder or leave it empty to manual archive"
                choice = show_input_popup(prompt, options)

                if choice in ['1', '2', '3']:
                    choice_int = int(choice) - 1
                    folder_path = top_matches.iloc[choice_int]['Path']
                elif choice == 'o':
                    folder_path = top_matches.iloc[0]['Path']
                    open_folder(folder_path)
                    print("Folder opened - archive manually")
                    exit()
                else:
                    folder_path = ''

        else:
            top_matches = find_top_matches(subject,sender,recipients)

            # print the three top options
            print("Top 3 matches:")
            for idx, row in top_matches.iterrows():
                print(f"{idx+1}: {row['Subject']} (Folder: {row['Path']})")

            # Get the user's choice from the popup
            options = [f"{idx + 1}: {row['Subject']} (Folder: {row['Path']})" for idx, row in
                       top_matches.iterrows()]
            prompt = "Enter the number of the chosen option (1/2/3) or 'o' to open the first folder or leave it empty to manual archive"
            choice = show_input_popup(prompt, options)

            if choice in ['1', '2', '3']:
                choice_int = int(choice) - 1
                folder_path = top_matches.iloc[choice_int]['Path']
            elif choice == 'o':
                folder_path = top_matches.iloc[0]['Path']
                open_folder(folder_path)
                print("Folder opened - archive manually")
                exit()
            else:
                folder_path = ''

if folder_path:

    # Calculate the next correlative number
    correlative_number = get_next_correlative_number(folder_path)
    print('correlative number = ' , correlative_number)

    # Save attachments
    save_attachments(email, folder_path, correlative_number)

    # Save email as .msg file
    save_email_as_msg(email, folder_path, correlative_number)

    # Archive the email (optional) uncomment the following line
    # archive_email(email)

else:
    print('No folder chosen. Archive manually')