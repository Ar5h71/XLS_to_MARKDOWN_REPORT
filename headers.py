import pandas as pd
import glob
import os


def gen_combined_sheet(sheets):
    all_sheets = []
    if len(sheets) == 1:
        for i in range(sheets[0].shape[0]):
            for j in range(sheets[0].shape[1]):
                sheets[0].iloc[i,j] = str(sheets[0].iloc[i,j])
        return sheets[0]
    all_sheets.append(pd.merge(sheets[0], sheets[1], on='_uuid'))
    if (len(sheets)) == 2:
        for i in range(all_sheets[0].shape[0]):
            for j in range(all_sheets[0].shape[1]):
                all_sheets[0].iloc[i,j] = str(all_sheets[0].iloc[i,j])
        return all_sheets[0]
    for j in range(len(sheets) - 2):
        all_sheets.append(pd.merge(all_sheets[0], sheets[j + 2], on='_uuid'))
    for k in range(len(all_sheets)):
        for i in range(all_sheets[k].shape[0]):
            for j in range(all_sheets[k].shape[1]):
                all_sheets[k].iloc[i, j] = str(all_sheets[k].iloc[i, j])
    return all_sheets[1:]


def rename_submission__uuid(sheets):
    for j in range(len(sheets)):
        sheets[j].rename(columns={'_submission__uuid': '_uuid'}, inplace=True)
    return sheets


def generate_headers(combined_sheet, spreadsheet_sheets):
    if len(spreadsheet_sheets)>2:
        for j in range(len(combined_sheet)):
            s = str(input("How do you want the output for "+ filename[i] + ' ' + spreadsheet_sheets[j+2] +" ? " +"Press 't' for table and 'l' for list: " + "\n"  ))
            dest.write(s + ',' + filename[i] + ',' + spreadsheet_sheets[j + 1] + ',')
            dim = combined_sheet[j].shape
            col_name = combined_sheet[j].columns
            index = 0
            cols = []
            for k in range(dim[1]):
                temp = 0
                if col_name[k][0] != '_':
                    l = len(col_name[k]) - 1
                    while col_name[k][l] != '>' and l > 0:
                        l = l - 1
                    for m in range(dim[0]):
                        if not 'nan' in combined_sheet[j].iloc[m,k]:
                            temp = 1
                            break
                    if temp:
                        if not l:
                            print(str(index+1) + ' ' + col_name[k])
                            cols.append(col_name[k])
                        else:
                            print(str(index+1) + ' ' + col_name[k][l + 1:])
                            cols.append(col_name[k][l + 1:])
                        index += 1
            col_index = []
            pqr = int(input("If you want the report in the above mentioned sequence, Press '0'" + "\n" + "If you want to give your own sequence, Press '1'"))
            if pqr:
                for k in range(index):
                    col_index.append(int(input("Enter Output You Want To See At Number " + str(k+1) + "\n" )))
                for m in range(len(cols)):
                    dest.write(cols[col_index[m] - 1] + ',')
                dest.write('\n')
            else:
                for m in range(index):
                    dest.write(cols[m] + ',')
                dest.write('\n')
    else:
        for n in range(combined_sheet.shape[0]):
            for o in range(combined_sheet.shape[1]):
                combined_sheet.iloc[n, o] = str(combined_sheet.iloc[n, o])
        s = str(input("How do you want the output for " + filename[i] + "? " + '\n' + "Press 't' for table and 'l' for list: " + '\n' ))
        dest.write(s + ',' + filename[i] + ',' + spreadsheet_sheets[len(spreadsheet_sheets) - 1] + ',')
        dim = combined_sheet.shape
        col_name = combined_sheet.columns
        index = 0
        cols = []
        for k in range(dim[1]):
            temp = 0
            if col_name[k][0] != '_':
                l = len(col_name[k]) - 1
                while col_name[k][l] != '>' and l > 0:
                    l = l - 1
                for m in range(dim[0]):
                    if not 'nan' in combined_sheet.iloc[m, k]:
                        temp = 1
                        break
                if temp:
                    if not l:
                        print(str(index+1) + ' ' + col_name[k])
                        cols.append(col_name[k])
                    else:
                        print(str(index+1) + ' ' + col_name[k][l + 1:])
                        cols.append(col_name[k][l + 1:])
                    index += 1
        col_index = []
        pqr = int(input("If you want the report in the above mentioned sequence, Press '0'" + "\n" + "If you want to give your own sequence, Press '1'"))
        if pqr:
            for k in range(index):
                col_index.append(int(input(
                    "Enter Output You Want To See At Number " + str(k + 1) +"\n")))
            for m in range(len(cols)):
                dest.write(cols[col_index[m] - 1] + ',')
            dest.write('\n')
        else:
            for m in range(index):
                    dest.write(cols[m] + ',')
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
    spreadsheet_sheets = spreadsheet.sheet_names  # gets the list of names of all the sheets in an excel file

    sheets = []
    for j in range(len(spreadsheet_sheets)):
        sheets.append(pd.read_excel(spreadsheet, spreadsheet_sheets[j]))

    sheets = rename_submission__uuid(sheets)

    combined_sheet = gen_combined_sheet(sheets)

    generate_headers(combined_sheet, spreadsheet_sheets)
