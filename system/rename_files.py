import os
import re

# source chatGPT 2024-05-25 > https://chatgpt.com/c/1049735e-6d2b-4603-8a4a-8f68b8de8e10

def rename_files_in_directory(directory):
    """
    Traverse through the given directory and its subdirectories,
    rename files by removing the 'v01', 'v02', ... 'vXX' pattern from the file names,
    and print the changes. Ensure no blank space before the extension.
    Skip renaming if the target file name already exists and recap the skipped files.

    :param directory: The path to the directory containing the files.
    """
    # Regular expression pattern to match 'vXX' or 'VXX' at the end of the filename (excluding the extension)
    pattern = re.compile(r'(.*?)([vV]\d{2})(\.\w+)$')
    changed_files_count = 0
    skipped_files = []

    for root, _, files in os.walk(directory):
        for file in files:
            # Match the file name with the pattern
            match = pattern.match(file)
            if match:
                old_name = os.path.join(root, file)

                # Strip any trailing whitespace from the filename part before the extension
                new_name = os.path.join(root, match.group(1).rstrip() + match.group(3))

                # Check if the new file name already exists
                if os.path.exists(new_name):
                    skipped_files.append(file)
                    print(f'Skipping file {file}')
                else:
                    # Rename the file
                    os.rename(old_name, new_name)
                    changed_files_count += 1

                    # Print the change
                    print(f'Renamed: {old_name} -> {new_name}')

    # Print the total number of files renamed
    print(f'\nTotal files renamed: {changed_files_count}')

    # Print the skipped files
    if skipped_files:
        print(f'\nSkipped files ({len(skipped_files)}):')
        for skipped_file in skipped_files:
            print(skipped_file)


# Request the directory path from the user
directory_path = input('Enter the directory path: ')
rename_files_in_directory(directory_path)
