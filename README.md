## Step 1: First copy the excel files, out of which you want to create reports, in the 'excel' folder
## Step 2: Run headers.py. It will extract all the column headers in a text file 'template.txt' in the template folder. Each line in the text file represents the columns for a file. The first element of each line is either 't' or 'l', denoting table and list respectively. The second element represents the file name.
## Step 3: Check whether you want a table or a list for a particular file and change the files element of each row accordingly in the text file. Also change the order of columns to your preference in the text file.
## Step 4: After you have decided the order of columns, run 'Report.py'. The report will be generated in a file  by the name 'report.md' in the 'markdown' folder.
