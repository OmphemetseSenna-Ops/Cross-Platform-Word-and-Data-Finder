# Automation_Audit
## File Sorting Automation for Auditors and CAATs
### Overview
This Python script automates the process of sorting downloaded files based on their contents, aiding auditors and CAATs (Computer-Assisted Audit Techniques) in efficiently identifying and selecting files for sampling. It supports various file types such as text files, Word documents, and Excel spreadsheets, allowing auditors to specify keywords for targeted file selection.

### Key Features
- File Type Selection: Choose from text files, Word documents, or Excel spreadsheets.
- Keyword Search: Specify keywords to search for within the selected files.
- Content-Based Sorting: Files containing the specified keywords are automatically copied to a designated folder.
- User-Friendly GUI: Built using Tkinter for easy folder selection and keyword entry.

### Use Case
Auditors often need to sift through numerous downloaded files to find relevant documents for audit sampling. This automation tool streamlines the process by automating file sorting based on predefined keywords, saving time and improving audit efficiency.

### How to Use
- Select Folder: Choose the directory containing the downloaded files.
- Choose File Type: Select the type of files to search (text, Word, or Excel).
- Enter Keyword: Input the keyword or phrase to search within the files.
- Run Automation: Click "Search and Copy" to initiate the search and sorting process.
- Review Results: View the number of files containing the keyword and confirm copied files in the designated folder.
- Dependencies
Python libraries: tkinter, pandas, python-docx