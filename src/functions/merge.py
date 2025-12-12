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

def csv_path() -> Path:
    """
    Obtain the file paths for the csv
    
    """

    while True:
        try:
            csv_path = input("Please provide csv file with the confirmation data: ").strip("\" ")
            csv_path = Path(csv_path)

            if not csv_path.is_file() or (csv_path.is_file() and csv_path.suffix != ".csv"):
                print("Please provide a valid csv file")
                continue
            # default result of Path is the cwd, therefore account for this so user is continuosly prompted if no input is provided
            elif str(csv_path) == ".":
                print("Please provide a valid csv file")
                continue
            print("CSV successfully provided")
            return csv_path
        except FileNotFoundError as e:
            print(e)
        except Exception as e:
            print(e)

def template_path() -> Path:
    """ 
    Obtain the file paths for the template
    
    """

    while True:
        try:
            template_path = input("Please provide your desired word template with fields matching the provided csv's headers: ").strip("\" ")
            template_path = Path(template_path)

            if not template_path.is_file() or (template_path.is_file and template_path.suffix != ".docx"):
                print("Please provide a valid word document")
                continue
            # default result of Path is the cwd, therefore account for this so user is continuosly prompted if no input is provided
            elif str(template_path) == ".":
                print("Please provide a valid word document")
                continue
            print("Word template successfully provided")
            return template_path
        except FileNotFoundError as e:
            print(e)
        except Exception as e:
            print(e)

def convert_word_to_pdf():
    """
    take a directory of word documents as input, convert each word document in that directory to pdf, output the pdfs in a new directory
    
    """

    while True:
        try:
            confirmation_path = input("Please provide your desired file path which contains confirmation drafts as word documents: ").strip("\" ")
            confirmation_path = Path(confirmation_path)

            if not confirmation_path.is_dir():
                print("Please provide a valid directory to confirmation drafts")
                continue
            # default result of Path is the cwd, therefore account for this so user is continuosly prompted if no input is provided
            elif str(confirmation_path) == ".":
                print("Please provide a valid directory to confirmation drafts")
                continue
            else:
                # track if other document in directory
                contains_only_word = True

                for file in confirmation_path.iterdir():
                    # check that input is file or directory and if file must be word
                    if not (file.is_file() or file.is_dir()) or (file.is_file() and file.suffix != ".docx"):
                        print("The directory provided contains documents which do not in '.docx' ")
                        print(file)
                        contains_only_word = False
                        break
                    else:
                        pass
                if contains_only_word == False:
                    continue
            print("Word confirmation path successfully provided")

            # the provided directory contains only word documents, create a directory for the pdf outputs and convert each word document to pdf at this location
            pdf_dir = confirmation_path / "pdf_versions"
            # exist okay = False will raise FileExistsError
            pdf_dir.mkdir(exist_ok=True)
            # batch convert all word to pdf
            try:
                convert(confirmation_path, pdf_dir)
            except Exception as e:
                print(e)
            print("Word documents successfully converted to pdf's")
            return 0
        except FileNotFoundError as e:
            print(e)
        except Exception as e:
            print(e)

def pdf_consolidation():
    """
    Take a directory which only contains pdf's, creates a new directory, for each pdf create a new pdf which consolidates all pdf's in the directory which have the same respondent name
    
    """

    while True:
        try:
            pdf_path = input("Please provide your desired file path which contains confirmation drafts as pdf documents: ").strip("\" ")
            pdf_path = Path(pdf_path)

            if not pdf_path.is_dir():
                print("Please provide a valid directory to pdf confirmation drafts")
                continue
            # default result of Path is the cwd, therefore account for this so user is continuosly prompted if no input is provided
            elif str(pdf_path) == ".":
                print("Please provide a valid directory to pdf confirmation drafts")
                continue
            else:
                # track if other type of document in the directory
                contains_only_pdf = True

                for file in pdf_path.iterdir():
                    # check that input is file or directory and if file must be pdf
                    if not (file.is_file() or file.is_dir()) or (file.is_file() and file.suffix != ".pdf"):
                        print("The directory provided contains documents which do not end in '.pdf' ")
                        print(file)
                        contains_only_pdf = False
                        break
                    else:
                        pass
                if contains_only_pdf == False:
                    continue
            print("PDF confirmation path successfully provided")

            # the provided directory contains only pdf documents, create a directory for the consolidated outputs and merge each pdf document by respondent at this location
            consolidate_dir = pdf_path / "consolidated_by_respondent"
            # exists okay = False will raise FileExistsError
            consolidate_dir.mkdir(exist_ok=True)

            # iterate through all the confirmations to compare against the suffixes
            for file in pdf_path.iterdir():
                if file.suffix == ".pdf":
                    # get file suffix 
                    file_suffix = str(file).rsplit("-", 1)

                    if len(file_suffix) > 1:
                        suffix = file_suffix[1]

                        merge_list = []
                        # the file suffix is compared to the suffix of every file in the directory and it matches, then it is appended to the list of documents to pass onto merging
                        for inner_file in pdf_path.iterdir():
                            if inner_file.suffix == ".pdf":
                                inner_suffix = str(inner_file).rsplit("-", 1)
                                inner_suffix = inner_suffix[1]

                                if inner_suffix == suffix:
                                    merge_list.append(inner_file)
                                else:
                                    pass

                    # if the document doesn't already exist at the output PATH, merge the documents in the list
                    if not (consolidate_dir / f"Confirmation_Consolidated_{file_suffix[1]}").exists() and len(merge_list) > 1:
                        consolidated_pdf_path_multiple = consolidate_dir / f"Confirmation_Consolidated_{file_suffix[1]}"
                        merge_pdfs(merge_list, consolidated_pdf_path_multiple)
                    elif len(merge_list) == 1:
                        file_name = str(file).rsplit("\\", 1)
                        file_name = file_name[1]
                        if not (consolidate_dir / file_name).exists():
                            consolidated_pdf_path_single = consolidate_dir / file_name
                            merge_pdfs(merge_list, consolidated_pdf_path_single)
                        else: 
                            pass
                    else:
                        pass
            print("PDF's successfully merged")
            return 0
        except FileNotFoundError as e:
            print(e)
        except Exception as e:
            print(e)

def merge_pdfs(pdf_list, output_path):
    # Create a PdfMerger object
    merger = PdfMerger()

    # Iterate over the list of PDF filenames and append them
    for pdf in pdf_list:
        merger.append(pdf)
    
    # Write the merged PDF to a specified path 
    merger.write(output_path)
    merger.close()

def pdf_suffix(pdf_path):
    suffixes = []
    
    try:
        for file in pdf_path.iterdir():
            parts = str(file).rsplit("-", 1)
            if len(parts) > 1:
                suffix = parts[1]

            if suffix not in suffixes:
                suffixes.append(suffix)
            else:
                pass

        sorted_suffixes = sorted(suffixes)
        return sorted_suffixes
    except Exception as e:
            print(e)

if __name__ == "__main__":
    main()