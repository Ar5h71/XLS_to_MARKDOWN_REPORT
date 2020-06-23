# XLS_to_MARKDOWN_REPORT
This repo helps to convert data exported from KoBoToolBox to Markdown format in tabular of list form and helps to automate the process of generating reports.
# download_excel.py
This script downloads excel files containing data from all the forms uploaded in Kobo Toolbox. 
# headers.py
This script extracts the useful headers from the excel files and writes the column headers, along with the filename and sheetname of excel file and the alphabets 't' or 'l' signifying table or list ,in a text file called 'template.txt' in the 'template' folder. It prompts the user to change the sequence of the column headers in the order they should appear in the final report. The user may also edit the 'template.txt' file later.
# Report.py
This script writes the data in a file 'report.md' in 'Markdown' folder followinf the format as provided in 'template.txt' file. It also sorts the data according to date in descending order and also converts the format of date in ddmmyyyy with a separator of user's choice.
# Usage Instructions
_Note:_ While exporting data manually off from KoBoToolBox, make sure that the group separators are changed to '>' and 'Include data from previous versions' is unticked.
1. Create the folders named 'excel', 'markdown' and 'template' in the directory where the python files 'headers.py' and 'Report.py' are stored.
2. Store the 'download_excel.py' python file in the 'excel' folder.
3. Run download_excel.py. It will save all the excel files in the 'excel' folder.
4. Run headers.py. It will extract all the column headers in a text file 'template.txt' in the template folder. Each line in the text file represents the columns for a file. The first element of each line is either 't' or 'l', denoting table and list respectively. The second element represents the file name. The third element is sheet name and the fourth element onwrds, column headers are written.
5. Check whether you want a table or a list for a particular file and change the files element of each row accordingly in the text file. Also change the order of columns to your preference in the text file.
6. After you have decided the order of columns, run 'Report.py'. The report will be generated in a file  by the name 'report.md' in the 'markdown' folder.
