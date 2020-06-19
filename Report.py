import pandas as pd
import os
import glob


def column_location(combined_sheet, indices):
    # determines the location of column headers according to the template file
    col_loc = []
    for i in range(len(indices)):
        for j in range(combined_sheet.shape[1]):
            if indices[i] in combined_sheet.columns[j]:
                col_loc.append(j)
    return col_loc


def date_sort_and_format(combined_sheet, length ,separator):
    all_sheets = []
    if length == 1:
        all_sheets.append(combined_sheet)
    else:
        for i in range(length):
            all_sheets.append(combined_sheet[i])

    for j in range(length):
        for i in range(all_sheets[j].shape[1]):
            if all_sheets[j].columns[i][0] != '_' and 'Date' in all_sheets[j].columns[i]:
                all_sheets[j][all_sheets[j].columns[i]] = pd.to_datetime(all_sheets[j][all_sheets[j].columns[i]])
                all_sheets[j] = all_sheets[j].sort_values(by=[all_sheets[j].columns[i]], ascending=False)
                break

    for j in range(length):
        for i in range(all_sheets[j].shape[1]):
            if all_sheets[j].columns[i][0] != '_' and 'Date' in all_sheets[j].columns[i]:
                all_sheets[j][all_sheets[j].columns[i]] = pd.to_datetime(all_sheets[j][all_sheets[j].columns[i]])
                all_sheets[j][all_sheets[j].columns[i]] = all_sheets[j][all_sheets[j].columns[i]].dt.strftime("%d" + separator + "%m" + separator + "%Y")

    if length == 1:
        return all_sheets[0]
    else:
        return all_sheets


def gen_combined_sheet(sheets):
    # merges the sheets on the basis of '_uuid' column that contains unique id for submissions
    all_sheets = []
    if len(sheets) == 1:
        return sheets[0]        # returns the first sheet if there is only one
    all_sheets.append(pd.merge(sheets[0], sheets[1], on='_uuid'))
    if (len(sheets)) == 2:
        return all_sheets[0]    # if there are only 2 sheets then merges both and returns the combined sheet
    for j in range(len(sheets) - 2):
        all_sheets.append(pd.merge(all_sheets[0], sheets[j + 2], on='_uuid'))
    return all_sheets[1:]       # if more than 2 sheets i.e. multi-category form then combines 3rd sheet onwards
                                # to 1st and 2nd sheet and returns a list of combined sheets


def write_table(combined_sheet, indices, f):
    for i in range(len(indices)):
        f.write('| ' + indices[i])
    f.write(' |\n')

    for i in range(len(indices)):
        f.write('|-----')
    f.write(' |\n')
    col_loc = column_location(combined_sheet, indices)

    for i in range(combined_sheet.shape[0]):
        for j in range(len(col_loc)):
            for k in range(len(combined_sheet.iloc[i, col_loc[j]])):
                if combined_sheet.iloc[i, col_loc[j]][k] == '\n':
                    combined_sheet.iloc[i, col_loc[j]] = combined_sheet.iloc[i, col_loc[j]][:k - 1] + ' <br> ' + \
                                                         combined_sheet.iloc[i, col_loc[j]][k + 1:]
            f.write('| ' + combined_sheet.iloc[i, col_loc[j]])
        f.write(' |\n')
    f.write('<br><br><br>\n')


def write_list(combined_sheet, indices, f):
    col_loc = column_location(combined_sheet, indices)

    for i in range(combined_sheet.shape[0]):
        f.write('- ')
        for j in range(len(col_loc) - 1):
            if 'nan' in combined_sheet.iloc[i, col_loc[j]]:
                continue
            else:
                f.write(combined_sheet.iloc[i, col_loc[j]] + ', ')
        if 'nan' not in combined_sheet.iloc[i, col_loc[len(col_loc) - 1]]:
            f.write(combined_sheet.iloc[i, col_loc[len(col_loc) - 1]])
        f.write('\n')
    f.write('<br><br><br>\n')


def rename_submission__uuid(sheets):
    # the unique id column name is '_uuid' in the first sheet and '_submission__uuid' in the subsequent sheets
    # this function renames '_submission__uuid' to '_uuid' so that the sheets can be merged
    for j in range(len(sheets)):
        sheets[j].rename(columns={'_submission__uuid': '_uuid'}, inplace=True)
    return sheets


def xls2md(excel_spreadsheet, file_name, template, f, separator):
    f.write('## **')

    for i in range(file_name.index('_-_')):       # determines the form name and converts it to all caps
        if file_name[i] !='_':
            f.write(file_name[i].upper())
        else:
            f.write(' ')

    spreadsheet_sheets = excel_spreadsheet.sheet_names
    sheet = []

    for m in range(len(spreadsheet_sheets)):      # stores all sheets in an excel file in a list 'sheet'
        sheet.append(pd.read_excel(excel_spreadsheet, spreadsheet_sheets[m]))
    indices = template[3:len(template) - 1]               # gets a list of names of column headers in a sheet

    sheet = rename_submission__uuid(sheet)
    combined_sheet = gen_combined_sheet(sheet)

    if len(sheet)<=2:               # merges the sheets in the excel file according to the column '_uuid'
        f.write('**\n')

        combined_sheet = date_sort_and_format(combined_sheet, 1, separator)

        for i in range(combined_sheet.shape[0]):
            for j in range(combined_sheet.shape[1]):
                combined_sheet.iloc[i,j] = str(combined_sheet.iloc[i,j])

        if template[0] == 't':
            write_table(combined_sheet, indices, f)
        elif template[0] == 'l':
            write_list(combined_sheet, indices, f)

    else:
        f.write(' - ' + template[2].replace('_', ' ').upper() + '**\n')

        combined_sheet = date_sort_and_format(combined_sheet, len(combined_sheet), separator)

        for i in range(len(combined_sheet)):
            for j in range(combined_sheet[i].shape[0]):
                for k in range(combined_sheet[i].shape[1]):
                    combined_sheet[i].iloc[j, k] = str(combined_sheet[i].iloc[j, k])

        for l in range(2, len(sheet)):
            if template[2] == spreadsheet_sheets[l]:
                break

        if template[0] == 't':
            write_table(combined_sheet[l-2], indices, f)
        elif template[0] == 'l':
            write_list(combined_sheet[l-2], indices, f)


cur_path = os.getcwd()  # gets the path where the python file is located
template_path = cur_path + r'\template'
path = cur_path + r'\excel'
template = open(template_path + '\\' + 'template.txt')
template_fields = template.readlines()

for i in range(len(template_fields)):
    template_fields[i] = template_fields[i].split(';')

files = [f for f in glob.glob(path + '**/*.xlsx', recursive=True)] # gets the paths for all the excel files stored in 'excel' folder
filename = []
lenpath = len(path) + 1
dest_path = os.getcwd() + '\\'
dest_path = dest_path + 'markdown\\'
dest_path = dest_path + 'report.md'  # destination file for markdown
sep = str(input("which separator do you want for dates? Please Enter\n"))
f = open(dest_path, 'w')

for file in files:
    filename.append(file[lenpath:len(file) - 5])          #gets the filenames of all the files in the 'excel' folder

for k in range(len(template_fields)):
    for i in range(len(filename)):
        if filename[i] in template_fields[k][1]:
            break
    spreadsheet = pd.ExcelFile(files[i])   #opens all the excel files one by one
    xls2md(spreadsheet, filename[i], template_fields[k],f, sep)     #the function that converts the excel spreadsheet to md

print("All Files have been converted to Markdown")
