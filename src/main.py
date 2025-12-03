from src.functions import merge
import argparse
import sys

def main():
    parser = cli()
    args = parser.parse_args()

    if args.mail_merge:
        merge.confirmation_drafts()
    elif args.convert_to_pdf:
        merge.convert_word_to_pdf()
    elif args.consolidate:
        merge.pdf_consolidation()
    else:
        sys.exit("Must provide and argument to the CLI")


def cli() -> argparse.ArgumentParser:
    """
    defines the cli options and returns to the main function

    """

    # intitialize the argumen parser
    parser = argparse.ArgumentParser(
        prog="confirmation_drafter", 
        description="CLI program which automates the drafting of confirmation via mailmerge, conversion to pdf, and consolidation of pdfs by respondent"
    )

    parser.add_argument("-mm", "--mail_merge", help="Creates word confirmation drafts from the provided word template and csv data", action="store_true")
    parser.add_argument("-pdf", "--convert_to_pdf", help="Converts all word documents in a provided directory into pdf documents", action="store_true")
    parser.add_argument("-c", "--consolidate", help="Consolidates all the pdfs in a provided directory into a single document", action="store_true")

    return parser

if __name__ == "__main__":
    main()

