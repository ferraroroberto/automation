# Automation for archiving email from Outlook to disk

Automate classifying email already in disk folders, and archive from outlook suggested based on the classification

Caution this library was entirely built with chatGPT, this readme file is a template that I'll complete

## Installation

 Python = 3.6 required

Other packages

- extract_msg
- os
- openpyxl 
- pandas
- re

To install the packages, use the following command

```bash
pip install extract_msg os openpyxl pandas re
```

### Quick Start

 Step by step

The first part takes the directory and the excel file with the database

```python
# Set the directory you want to search in
dir_path = r'EonedriveDocumentosRoberto'

# Load the existing Excel file as a DataFrame
excel_path = r'EonedriveDocumentosRobertoprojectsautomationemail_managementemail_management.xlsx'
try
    df_existing = pd.read_excel(excel_path)
except FileNotFoundError
    df_existing = pd.DataFrame(columns=[Email Subject, Path, Sender, Recipients, Archive, Date Added])
```

- Get the column widths from the existing Excel file
- Create a list to store the file names and paths
- Initialize a counter for processed emails
- Iterate over all the files in the directory
- Check if the file has a .msg extension
- Get the full path to the file
- Check if the file is already in the existing DataFrame

The core of the code is this part, where the program cleans the name of the email message and extracts, sender and recipients

```python
            # Remove leading number sequence and dash from file name
            file_name = re.sub(r^d+s-s, , file)

            # Remove re or rv prefix and any leading whitespace from email subject
            subject = re.sub(r^s(rervfw[d])ss, , file_name, flags=re.IGNORECASE)

            # Extract sender and recipient information from the .msg file, with error control
            try
                with extract_msg.Message(file_path) as msg
                    sender = msg.sender
                    recipients = msg.to
            except InvalidFileFormatError
                print(fInvalidFileFormatError Skipping file {file_path})
                continue
                
            # Add the modified subject, path, sender, recipients, and other information to the list
            file_list.append((subject.strip(), subdir, sender, recipients, None, pd.Timestamp.now()))

            # Increment the counter
            processed_emails += 1

            # Pause and ask for an Enter keypress after 1,000 emails processed
            if processed_emails % 1000 == 0
                input(fProcessed {processed_emails} emails. Press 'Enter' to continue...)
```

Then

- Create a DataFrame from the file list
- Rename the Archive, Date Added, Sender and Recipients columns in the existing DataFrame 
- Merge the existing and new DataFrames on the Email Subject and Path columns, keeping all rows
- Combine the Sender Existing and Sender New columns into a single Sender column
- Combine the Recipients Existing and Recipients New columns into a single Recipients column
- Combine the Archive Existing and Archive New columns into a single Archive column
- Combine the Date Added Existing and Date Added New columns into a single Date Added column
- Remove duplicate email subjects and paths
- Export the updated database to an Excel file
- Apply the column widths to the final Excel file

## Usage

 Example of usage

- enter the folder we want to scan
- execute
- check

## Documentation

For comprehensive documentation, including available methods and parameters, here are the chatGPT prompts Will work only for me, since I need to login.

[prompt 1 - initial prompt](httpschat.openai.comce39d4cde-d4ef-434a-b32d-5b67ce52b72b)

[prompt 2 - column width, keyboard input, extra information](httpschat.openai.comcc6c3d27c-779a-45ba-9694-8e7b8605c057)

[prompt 3 - error control](httpschat.openai.comcf6d2c233-6a5f-4783-a457-72803280cd40)

[prompt 4 - managing outlook](httpschat.openai.comc44fb50fc-905f-4be2-bcae-a370fc5c6d75)

## Disclaimer

Text of the disclaimer.

## Contributing

We welcome contributions! [troubleshooting](#troubleshooting-errors)

## Development

### Troubleshooting errors

#### Error 1

Sample error..

Known reasons include

- reason one
- reson two
