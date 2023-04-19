from fuzzywuzzy import fuzz
import pandas as pd
import re
import win32com.client
import tkinter as tk
from tkinter import messagebox
import os
import string
from pathlib import Path

# Function to read the parameters from the txt file
def read_params_from_txt_file(file_path):
    params = {}
    with open(file_path, 'r') as f:
        for line in f:
            if line.strip():
                key, value = line.strip().split(" = ", 1)
                params[key.strip()] = value.strip()
    return params

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

    # Calculate the total score by giving equal weight to the subject similarity and the sender/receiver similarity
    df['Total_Score'] = df['Subject_Score'] * 0.5 + df['Send_Recv_Score'] * 0.5

    # Get the top 3 matches based on the total score
    top_matches = df.nlargest(3, 'Total_Score')

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

    tk.Label(window, text=prompt).pack()

    yes_button = tk.Button(window, text="Yes, archive", command=on_yes)
    yes_button.pack()

    no_button = tk.Button(window, text="No", command=on_no)
    no_button.pack()

    open_folder_button = tk.Button(window, text="Open Folder", command=on_open_folder)
    open_folder_button.pack()

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

    tk.Label(window, text=prompt, anchor=tk.W).pack(anchor=tk.W)

    # Display the folder options
    for option in options:
        folder_path = option.split("(Folder: ")[1].rstrip(")")
        tk.Label(window, text=folder_path, anchor=tk.W).pack(anchor=tk.W)

    input_entry = tk.Entry(window)
    input_entry.pack()

    # Set focus to the input entry
    input_entry.focus_set()

    submit_button = tk.Button(window, text="Submit", command=on_submit)
    submit_button.pack()

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

import string

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

# Load the Excel file
excel_path = params['excel_path']
df = pd.read_excel(excel_path)

# Evaluate the email and decide
email = get_selected_email()

if email:

    # Remove "re" or "rv" prefix and any leading/trailing whitespace or ".msg" suffix from email subject
    subject = re.sub(r"^\s*(re|rv|fw[d]*)\s*:?\s*|\s*\.msg$", "", email.Subject, flags=re.IGNORECASE)

    sender = email.SenderEmailAddress
    recipients = ';'.join([r.Address for r in email.Recipients])

    print('subject = ' + subject + ' sender = ' + sender + ' - recipients = ' + recipients)

    match = search_email(subject, sender, recipients)
    if not match.empty:
        print(f"Perfect match by subject, sender, recipient found. The folder path is: {folder_path}")

        # Show a Yes/No popup asking if the user wants to proceed with the perfect match
        prompt = f"Perfect match by subject, sender, recipient found. The folder path is: {match.iloc[0]['Path']}. Do you want to proceed?"
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
            prompt = f"Subject match. The folder path is: {match.iloc[0]['Path']}. Do you want to proceed?"
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