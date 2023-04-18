from fuzzywuzzy import fuzz #requires 'pip install fuzzywuzzy python-Levenshtein'
import pandas as pd #requires 'pip install pandas openpyxl'
import re
import win32com.client # requires 'pip install pywin32'
import tkinter as tk
from tkinter import messagebox

# Load the Excel file
excel_path = r'E:\onedrive\Documentos\Roberto\projects\automation\email-automation\email-archive\email-archive.xlsx'
df = pd.read_excel(excel_path)

# Function to search for an email in the Excel file
def search_email(subject, sender, recipients):
    filtered_df = df.loc[(df['Subject'] == subject) & (df['Sender'] == sender) & (df['Recipients'] == recipients)]
    return filtered_df

# Function to search for an email in the Excel file, only for subject
def search_email_subject(subject):
    filtered_df = df.loc[(df['Subject'] == subject)]
    return filtered_df

# Function to find the top 3 most likely subjects and corresponding folders
def find_top_matches(subject):
    df['Subject_Score'] = df['Subject'].apply(lambda x: fuzz.token_set_ratio(subject, x))
    top_matches = df.nlargest(3, 'Subject_Score')
    return top_matches

# Function to get the selected email from Outlook
def get_selected_email():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    explorer = outlook.Application.ActiveExplorer()
    selection = explorer.Selection

    if len(selection) == 0:
        print("No email is selected.")
        return None

    return selection.Item(1)

# Function to create a popup window with radio buttons
def show_choice_popup(options):
    def on_submit():
        nonlocal selected_option
        selected_option.set(var.get())
        window.destroy()

    window = tk.Tk()
    window.title("Choose a folder")

    var = tk.IntVar()
    for i, option in enumerate(options):
        tk.Radiobutton(window, text=option, variable=var, value=i).pack(anchor=tk.W)

    tk.Button(window, text="Submit", command=on_submit).pack()

    selected_option = tk.IntVar(value=-1)
    window.mainloop()
    return selected_option.get()

#execution
email = get_selected_email()

if email:

    # Remove "re" or "rv" prefix and any leading/trailing whitespace or ".msg" suffix from email subject
    subject = re.sub(r"^\s*(re|rv|fw[d]*)\s*:?\s*|\s*\.msg$", "", email.Subject, flags=re.IGNORECASE)

    sender = email.SenderEmailAddress
    recipients = ';'.join([r.Address for r in email.Recipients])

    print('subject = ' + subject + ' sender = ' + sender + ' - recipients = ' + recipients)

    match = search_email(subject, sender, recipients)
    if not match.empty:
        folder_path = match.iloc[0]['Path']
        print(f"Perfect match by subject, sender, recipient found. The folder path is: {folder_path}")
    else:

        match = search_email_subject(subject)
        if not match.empty:
            folder_path = match.iloc[0]['Path']
            print(f"Match by subject found. The folder path is: {folder_path}")
        else:
            top_matches = find_top_matches(subject)

            # print the three top options
            print("Top 3 matches:")
            for idx, row in top_matches.iterrows():
                print(f"{idx+1}: {row['Subject']} (Folder: {row['Path']})")

            # popup with the top options
            options = [f"{idx + 1}: {row['Subject']} (Folder: {row['Path']})" for idx, row in top_matches.iterrows()]
            choice = show_choice_popup(options)

            if choice in [0, 1, 2]:
                chosen_folder = top_matches.iloc[choice]['Path']
                messagebox.showinfo("Chosen folder", f"Chosen folder: {chosen_folder}")
            else:
                messagebox.showinfo("No folder chosen", "No folder chosen.")
