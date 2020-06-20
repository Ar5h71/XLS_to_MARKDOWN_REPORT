# This file generates the template for the report and writes the template in 'template.txt' file in template folder

# The 'template.txt' file can also later be edited to remove some fields or change the order or type of format for a particular form

# Each line in the template file contains template for each form.

# The first element of each individual line tells whether the data will be represented in tabular or list form.
# The second element contains the filename for that particular form.
# The third element contains the sheetname. It is useful in case of forms where set of questions vary according to catgery selected
# fourth element onwards, list of all column headers is written


import pandas as pd
import glob
import os


def gen_combined_sheet(sheets):
    """
        Returns a list of DataFrames which can be worked upon to generate reports.

        This function takes a list of dataframes which are all the sheets in an excel file and merges the relevant sheets
        according to the unique id column '_uuid'.

            Parameters:
                sheets (list of DataFrames): A list of pandas DataFrames. All the sheets of the excel file opened as DataFrames and stored in a list.

            Returns:
                all_sheets (List of DataFrames):A list of DataFrames after performing the merging operation.
    """
    all_sheets = []
    if len(sheets) == 1:
        for i in range(sheets[0].shape[0]):
            for j in range(sheets[0].shape[1]):
                sheets[0].iloc[i,j] = str(sheets[0].iloc[i,j])
        return sheets[0]        # returns the first sheet if there is only one
    all_sheets.append(pd.merge(sheets[0], sheets[1], on='_uuid'))
    if (len(sheets)) == 2:
        for i in range(all_sheets[0].shape[0]):
            for j in range(all_sheets[0].shape[1]):
                all_sheets[0].iloc[i,j] = str(all_sheets[0].iloc[i,j])
        return all_sheets[0]    # if there are only 2 sheets then merges both and returns the combined sheet
    for j in range(len(sheets) - 2):
        all_sheets.append(pd.merge(all_sheets[0], sheets[j + 2], on='_uuid'))
    for k in range(len(all_sheets)):
        for i in range(all_sheets[k].shape[0]):
            for j in range(all_sheets[k].shape[1]):
                all_sheets[k].iloc[i, j] = str(all_sheets[k].iloc[i, j])
    return all_sheets[1:]       # if more than 2 sheets i.e. multi-category form then combines 3rd sheet onwards
                                # to 1st and 2nd sheet and returns a list of combined sheets


def rename_submission__uuid(sheets):
    """
       Renames the '_submission__uuid' column to '_uuid'.

       The unique id column name is '_uuid' in the first sheet and '_submission__uuid' in the subsequent sheets.
       This function renames '_submission__uuid' to '_uuid' so that the sheets can be merged.

           Parameters:
               sheets (list of DataFrames): a list of DataFrames containing data of all the sheets in an excel file.

           Returns:
               sheets (list of DataFrames): list of DataFrames with '_submission__uuid' renamed as '_uuid'
    """
    for j in range(len(sheets)):
        sheets[j].rename(columns={'_submission__uuid': '_uuid'}, inplace=True)
    return sheets


def generate_headers(sheets, sheetnames, file_name, dest):
    """
    Writes all the column headers in 'template.txt'.

        Parameters:
            sheets (list of dataframes): All sheets of the excel file opened as dataframes and stored in a list.
            sheetnames (list of strings): contains names of all the sheets in the excel file.
            file_name (string): name of the file.
            dest (file object): file object to write the column headers in 'template.txt'.
    """
    combined_sheet = gen_combined_sheet((sheets))
    if len(sheets)>2:  # if the form is multi-category form
        all_sheets = combined_sheet
    else:
        all_sheets = []
        all_sheets.append(combined_sheet)

    for j in range(len(all_sheets)):
        if len(sheets)>2:   # for multi-category forms that have different sets of repeating questions.
            s = str(input("How do you want the output for " + file_name + ' ' + sheetnames[j + 2] + " ? " + "Press 't' for table and 'l' for list: " + "\n"))
            dest.write(s + ';' + filename[i] + ';' + sheetnames[j + 2] + ';')
        else:       # for simple forms or forms having only one set of repeating questions.
            s = str(input("How do you want the output for " + file_name + ' ' +  " ? " + "Press 't' for table and 'l' for list: " + "\n"))
            dest.write(s + ';' + filename[i] +sheetnames[len(sheetnames) - 1]+ ';' +  ';')
        dim = all_sheets[j].shape
        col_name = all_sheets[j].columns
        index = 0
        cols = []
        for k in range(dim[1]):  # prints all the column headers in the console
            temp = 0
            if col_name[k][0] != '_':
                l = len(col_name[k]) - 1
                while col_name[k][l] != '>' and l > 0:
                    l = l - 1
                for m in range(dim[0]):
                    if not 'nan' in all_sheets[j].iloc[m, k]:
                        temp = 1
                        break
                if temp:
                    if not l:
                        print(str(index + 1) + ' ' + col_name[k])
                        cols.append(col_name[k])
                    else:
                        print(str(index + 1) + ' ' + col_name[k][l + 1:])
                        cols.append(col_name[k][l + 1:])
                    index += 1
        col_index = []
        pqr = int(input(
            "If you want the report in the above mentioned sequence, Press '0'" + "\n" + "If you want to give your own sequence, Press '1'\n"))
        if pqr:
            for k in range(index):  # sorts the column according to the sequence given bu the user
                col_index.append(int(input("Enter Output You Want To See At Number " + str(k + 1) + "\n")))
            for m in range(len(cols)):
                dest.write(cols[col_index[m] - 1] + ';')
            dest.write('\n')
        else:
            for m in range(index):
                dest.write(cols[m] + ';')
            dest.write('\n')


cur_path = os.getcwd()  # gets the path where the python file is located
path = cur_path + r'\excel'
files = [f for f in glob.glob(path + '**/*.xlsx', recursive=True)]  # gets the paths for all the excel files stored in 'excel' folder
filename = []
lenpath = len(path) + 1
dest_path = os.getcwd() + '\\'
dest_path = dest_path + 'template\\'
dest_path = dest_path + 'template' + '.txt'  # destination file for markdown
dest = open(dest_path, "w")

for f in files:
    filename.append(f[lenpath:len(f) - 5])  # gets the filenames of all the files in the 'excel' folder

for i in range(len(filename)):
    spreadsheet = pd.ExcelFile(files[i])  # opens all the excel files one by one
    sheetnames = spreadsheet.sheet_names  # gets the list of names of all the sheets in an excel file

    sheets = []
    for j in range(len(sheetnames)):
        sheets.append(pd.read_excel(spreadsheet, sheetnames[j]))

    sheets = rename_submission__uuid(sheets)

    generate_headers(sheets, sheetnames, filename[i], dest)

