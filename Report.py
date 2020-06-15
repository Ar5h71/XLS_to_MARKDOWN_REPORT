import pandas as pd
import os
import glob


def location(combined_sheet, indices):
    col_loc = []
    for i in range(len(indices)):
        for j in range(combined_sheet.shape[1]):
            if indices[i] in combined_sheet.columns[j]:
                col_loc.append(j)
    return col_loc


def xls2md_table(excel_spreadsheet, file_name, cols, f):
    f.write('## ')
    for i in range(file_name.index('_-_')):
        if file_name[i] !='_':
            f.write(file_name[i].upper())
        else:
            f.write(' ')
    spreadsheet_sheets = excel_spreadsheet.sheet_names
    sheet = []
    for m in range(len(spreadsheet_sheets)):
        sheet.append(pd.read_excel(excel_spreadsheet, spreadsheet_sheets[m]))

    indices = cols[3:len(cols) - 1]  # gets a list of names of column headers in a sheet
    for i in range(1, len(sheet)):
        sheet[i].rename(columns={'_submission__uuid':'_uuid'}, inplace = True)
    if len(spreadsheet_sheets)<=2:
        f.write('\n')
        combined_sheet = pd.merge(sheet[0], sheet[1], on = '_uuid')
        for i in range(2, len(sheet)):
            combined_sheet = pd.merge(combined_sheet, sheet[i], on = '_uuid')
        for i in range(combined_sheet.shape[0]):
            for j in range(combined_sheet.shape[1]):
                combined_sheet.iloc[i,j] = str(combined_sheet.iloc[i,j])
        for i in range(len(indices)):
            f.write('| ' + indices[i])
        f.write(' |\n')
        for i in range(len(indices)):
            f.write('|-----')
        f.write(' |\n')
        col_loc = location(combined_sheet, indices)
        for i in range(combined_sheet.shape[0]):
            for j in range(len(col_loc)):
                for k in range(len(combined_sheet.iloc[i, col_loc[j]])):
                    if combined_sheet.iloc[i, col_loc[j]][k] == '\n':
                        combined_sheet.iloc[i, col_loc[j]] = combined_sheet.iloc[i, col_loc[j]][:k-1] + ' <br> ' + combined_sheet.iloc[i, col_loc[j]][k+1:]
                f.write('| ' + combined_sheet.iloc[i,col_loc[j]])
            f.write(' |\n')
        f.write('<br><br><br>\n')
    else:
        combined_sheet = []
        f.write(' - ' + cols[2].replace('_', ' ').upper() + '\n')
        temp = pd.merge(sheet[0], sheet[1], on = '_uuid')
        for i in range(len(sheet)-2):
            combined_sheet.append(pd.merge(temp, sheet[i+2], on = '_uuid'))
        for i in range(len(combined_sheet)):
            for j in range(combined_sheet[i].shape[0]):
                for k in range(combined_sheet[i].shape[1]):
                    combined_sheet[i].iloc[j,k] = str(combined_sheet[i].iloc[j,k])
        for i in range(len(indices)):
            f.write('| ' + indices[i])
        f.write(' |\n')
        for i in range(len(indices)):
            f.write('|-----')
        f.write(' |\n')

        for l in range(2, len(sheet)):
            if cols[2] == spreadsheet_sheets[l]:
                break
        col_loc = location(combined_sheet[l-2], indices)
        for i in range(combined_sheet[l-2].shape[0]):
            for j in range(len(col_loc)):
                for k in range(len(combined_sheet[l-2].iloc[i, col_loc[j]])):
                    if combined_sheet[l-2].iloc[i, col_loc[j]][k] == '\n':
                        combined_sheet[l-2].iloc[i, col_loc[j]] = combined_sheet[l-2].iloc[i, col_loc[j]][:k - 1] + ' <br> ' + combined_sheet[l-2].iloc[i, col_loc[j]][k + 1:]
                f.write('| ' + combined_sheet[l-2].iloc[i, col_loc[j]])
            f.write(' |\n')
        f.write('<br><br><br>\n')


def xls2md_list(excel_spreadsheet, file_name, cols, f):
    f.write('## ')
    for i in range(file_name.index('_-_')):
        if file_name[i] != '_':
            f.write(file_name[i].upper())
        else:
            f.write(' ')
    spreadsheet_sheets = excel_spreadsheet.sheet_names
    sheet = []
    for m in range(len(spreadsheet_sheets)):
        sheet.append(pd.read_excel(excel_spreadsheet, spreadsheet_sheets[m]))

    indices = cols[3:len(cols) - 1]  # gets a list of names of column headers in a sheet
    for i in range(1, len(sheet)):
        sheet[i].rename(columns={'_submission__uuid':'_uuid'}, inplace = True)
    if len(spreadsheet_sheets)<=2:
        f.write('\n')
        combined_sheet = sheet[0]
        for i in range(1, len(sheet)):
            combined_sheet = pd.merge(combined_sheet, sheet[i], on = '_uuid')
        for i in range(combined_sheet.shape[0]):
            for j in range(combined_sheet.shape[1]):
                combined_sheet.iloc[i,j] = str(combined_sheet.iloc[i,j])
        col_loc = location(combined_sheet, indices)

        for i in range(combined_sheet.shape[0]):
            f.write('- ')
            for j in range(len(col_loc)-1):
                f.write(combined_sheet.iloc[i,col_loc[j]] + ', ')
            f.write(combined_sheet.iloc[i,col_loc[len(col_loc)-1]])
            f.write('\n')
        f.write('<br><br><br>\n')
    else:
        f.write(' - ' + cols[2].replace('_', ' ').upper() + '\n')
        combined_sheet = []
        temp = pd.merge(sheet[0], sheet[1], on='_uuid')
        for i in range(len(sheet) - 2):
            combined_sheet.append(pd.merge(temp, sheet[i + 2], on='_uuid'))
            print(combined_sheet[i].shape)
        for i in range(len(combined_sheet)):
            for j in range(combined_sheet[i].shape[0]):
                for k in range(combined_sheet[i].shape[1]):
                    combined_sheet[i].iloc[j, k] = str(combined_sheet[i].iloc[j, k])

        for l in range(2, len(sheet)):
            if cols[2] == spreadsheet_sheets[l]:
                break
        col_loc = location(combined_sheet[l - 2], indices)
        for i in range(combined_sheet[l-2].shape[0]):
            f.write('- ')
            for j in range(len(col_loc) - 1):
                f.write(combined_sheet[l-2].iloc[i, col_loc[j]] + ', ')
            f.write(combined_sheet[l-2].iloc[i, col_loc[len(col_loc) - 1]])
            f.write('\n')
        f.write('<br><br><br>\n')


cur_path = os.getcwd()  # gets the path where the python file is located
template_path = cur_path + r'\template'
path = cur_path + r'\excel'
template = open(template_path + '\\' + 'template.txt')
template_fields = template.readlines()
for i in range(len(template_fields)):
    template_fields[i] = template_fields[i].split(',')
files = [f for f in glob.glob(path + '**/*.xlsx', recursive=True)] # gets the paths for all the excel files stored in 'excel' folder
filename = []
lenpath = len(path) + 1
dest_path = os.getcwd() + '\\'
dest_path = dest_path + 'markdown\\'
dest_path = dest_path + 'report.md'  # destination file for markdown
f = open(dest_path, 'w')

for file in files:
    filename.append(file[lenpath:len(file) - 5])          #gets the filenames of all the files in the 'excel' folder


for k in range(len(template_fields)):
    for i in range(len(filename)):
        if filename[i] in template_fields[k][1]:
            break
    spreadsheet = pd.ExcelFile(files[i]) #opens all the excel files one by one
    if template_fields[k][0] == 't':
        xls2md_table(spreadsheet, filename[i], template_fields[k],f)     #the function that converts the excel spreadsheet to md
    elif template_fields[k][0] == 'l':
        xls2md_list(spreadsheet, filename[i], template_fields[k],f)
print("All Files have been converted to Markdown")
