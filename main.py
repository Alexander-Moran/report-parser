import os
import sys
from docx import Document
import reportParser

print("Running IIC Hacker Hub")
rp = reportParser.reportParser()
directory = os.path.join(os.pardir, 'reports')

def loading_bar(current, total, bar_length=40):
    percent = (current / total) * 100
    hashes = '█' * int(current / total * bar_length)
    spaces = ' ' * (bar_length - len(hashes))
    sys.stdout.write(f"\r[{hashes}{spaces}] {percent:.2f}%")
    sys.stdout.flush()

if os.path.exists(directory) and os.path.isdir(directory):
    files = [f for f in os.listdir(directory) if os.path.isfile(os.path.join(directory, f))]
    total_files = len(files)

    # Iterate through the files in the directory
    for idx, filename in enumerate(files, start=1):
        file_path = os.path.join(directory, filename)
        print(f"\rProcessing: {filename}", end="")

        document = Document(file_path)
        rp.report_parser(document)  # Assuming this is where you're parsing the report

        # Update the loading bar after processing each file
        loading_bar(idx, total_files)

    print()

else:
    print(f"Directory '{directory}' does not exist.")

rp.csv_export()


