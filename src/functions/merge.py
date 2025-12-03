from mailmerge import MailMerge
from pathlib import Path
from docx2pdf import convert
from PyPDF2 import PdfMerger
import csv

def main():
    pdf_consolidation()

def confirmation_drafts() -> None:
    """
    Takes a csv file with the confirmation data and maps to a word template which has been prepared with merge fields to draft the confirmations
    
    """

output_dir = output_path()
csv_file = csv_path()
template = template_path()

# obtain the merge fields so that the fields can be supplied via kwargs unpacking
merge_fields = obtain_merge_fields(template)

# ensure that the encoding parameter is set, otherwise the respondent column will be truncated
with open(csv_file, newline="", encoding="utf-8-sig") as confirmation_data:
    reader = csv.DictReader(confirmation_data)

    for row in reader:
        # initialize a dictionary based off of the merge fields of the template to unpack and map the csv data to the template
        merge_dict = {field.strip(): row.get(field.strip(), '') for field in merge_fields}

        with MailMerge(template) as document:
            document.merge(**merge_dict)
            document.write(f"{output_dir}/Confirmation-Draft-{row['entity']}-{row['respondent']}.docx")

def obtain_merge_fields(template):
    with MailMerge(template) as document:
        merge_fields = document.get_merge_fields()
    return merge_fields

def output_path() -> Path:
    """
    Obtain the output location for the user

    """

    while True:
        try:
            output_path = input("Please provide your desired file path to store the confirmation draft output: ").strip("\" ")
            output_path = Path(output_path)

            if not output_path.is_dir():
                print("Please provide a valid directory to output the confirmation drafts")
                continue
            elif str(output_path) == ".":
                print("Please provide a valid directory to output the confirmation drafts")
                continue
            print("Output path successfully provided")
            return output_path
        except FileNotFoundError as e:
            print(e)
        except Exception as e:
            print(e)



