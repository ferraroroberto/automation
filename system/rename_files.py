import logging
import os
import re

logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
log = logging.getLogger(__name__)

def rename_files_in_directory(directory: str) -> None:
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
                    log.info("Skipping file %s", file)
                else:
                    # Rename the file
                    os.rename(old_name, new_name)
                    changed_files_count += 1
                    log.info("Renamed: %s -> %s", old_name, new_name)

    log.info("Total files renamed: %d", changed_files_count)

    if skipped_files:
        log.info("Skipped files (%d):", len(skipped_files))
        for skipped_file in skipped_files:
            log.info("  %s", skipped_file)


if __name__ == "__main__":
    # Request the directory path from the user
    directory_path = input('Enter the directory path: ')
    rename_files_in_directory(directory_path)
