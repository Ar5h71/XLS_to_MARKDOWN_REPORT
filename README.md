# While downloading reports off from kobo change the group separator for excel file from '/' to '>'

## Step 1: Create the folders named 'excel', 'markdown' and 'template' in the directory where the python files 'headers.py' and 'Report.py' are stored.
## Step 2: Copy the excel files, out of which you want to create reports, in the 'excel' folder
## Step 3: Run headers.py. It will extract all the column headers in a text file 'template.txt' in the template folder. Each line in the text file represents the columns for a file. The first element of each line is either 't' or 'l', denoting table and list respectively. The second element represents the file name.
## Step 4: Check whether you want a table or a list for a particular file and change the files element of each row accordingly in the text file. Also change the order of columns to your preference in the text file.
## Step 5: After you have decided the order of columns, run 'Report.py'. The report will be generated in a file  by the name 'report.md' in the 'markdown' folder.
