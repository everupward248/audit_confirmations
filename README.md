# Audit Confirmation Drafter

## Backstory
As an auditor, one of the tasks which I periodically have to do is the drafting of external confirmation requests to independent 3rd parties. The first time I had to do this, it was one of the
most tedious and mindnumbing tasks I've had to do. So much time spent copying over information from spreadsheets into word documents. So much room for human error and if some kind of systematic error
was made then the whole batch would have to be redone. My second time around, I was set on never doing a manual confirmation draft again. This script has personally saved me so much headaches and time 
and has made this unpleasant process into something manageable and easy. 

## Description
This program is a CLI used in the drafting of external confirmation requests to be used for gathering evidence in an audit. There are three main components: mail merge, convert to pdf, and consolidating the pdf's by respondent.
1. Mail merge is used here to take a word template and map csv data to the merge fields in the word document
2. The file path containing the word drafts is passed as argument and all files in that directory are converted to pdf.
3. The file path containing the pdf drafts is passed as argument and all files which end in the same respondent suffix are merged into one pdf file

### Dependencies
The mail merge is performed using `mailmerge`, the conversion to pdf using `docx2pdf`, and the consolidation of the pdf files using `PyPDF2`
