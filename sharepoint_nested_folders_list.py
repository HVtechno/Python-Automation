import os
import re
import csv

# Set the path to the directory
dir_path = "(Your Local drive path)"

# Create an empty list to store folder names
folder_list = []

# Define a regular expression to match non-alphanumeric characters
regex = re.compile(r"[~\!\@\#\$\%\^\*\"\'\?\,\.\|\_\{\}\[\]\/\–\-]")

# Walk through the directory and its subdirectories
for root, dirs, files in os.walk(dir_path):
    for dir_name in dirs:
        # Remove special characters from folder name
        clean_dir_name = regex.sub("", dir_name)
        # Add cleaned folder name to list
        folder_list.append(os.path.join(root, clean_dir_name))

# Set the path to the output CSV file
output_file = "Example.csv"

# Write the list of folder names to the output CSV file
with open(output_file, "w", newline="") as f:
    writer = csv.writer(f)
    writer.writerow(["Folder Name"])
    for folder in folder_list:
        writer.writerow([folder])

# Print the list of folder names
print(folder_list)