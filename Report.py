import pandas as pd
import os
import glob


def column_location(combined_sheet, indices):    #determines the location of column headers according to the template file
    col_loc = []
    for i in range(len(indices)):
        for j in range(combined_sheet.shape[1]):
            if indices[i] in combined_sheet.columns[j]:
                col_loc.append(j)
    return col_loc


def gen_combined_sheet(sheets):     #combines the sheets in the excel file on the basis of unique id column '_uuid'
    all_sheets = []                 
    if len(sheets) == 1:            #if there is only one sheet then the sheet itself is returned.
        for i in range(sheets[0].shape[0]):
            for j in range(sheets[0].shape[1]):
                sheets[0].iloc[i,j] = str(sheets[0].iloc[i,j])
        return sheets[0]
    all_sheets.append(pd.merge(sheets[0], sheets[1], on='_uuid'))
    if (len(sheets)) == 2:          #if there are two sheets then they both are merged and then returned
        for i in range(all_sheets[0].shape[0]):
            for j in range(all_sheets[0].shape[1]):
                all_sheets[0].iloc[i,j] = str(all_sheets[0].iloc[i,j])
        return all_sheets[0]
    for j in range(len(sheets) - 2):   #if there are more than two sheets then first sheet and second sheet are combined and the rest of the sheets are merged with these two one by one and returned as a list.
        all_sheets.append(pd.merge(all_sheets[0], sheets[j + 2], on='_uuid'))
    for k in range(len(all_sheets)):
        for i in range(all_sheets[k].shape[0]):
            for j in range(all_sheets[k].shape[1]):
                all_sheets[k].iloc[i, j] = str(all_sheets[k].iloc[i, j])
    return all_sheets[1:]


def rename_submission__uuid(sheets):         #the unique id column is '_uuid' in the first sheet and '_submission__uuid' in the rest.
    for j in range(len(sheets)):             #this function renames the '_submission__uuid' to '_uuid' so that the sheets can be merged
        sheets[j].rename(columns={'_submission__uuid': '_uuid'}, inplace=True)
    return sheets


def xls2md_list(excel_spreadsheet, file_name, cols, f, separator):  #writes data as a list for a specific form
    f.write('## **')

    for i in range(file_name.index('_-_')):             #Extracts the form name from file name and replaces '_' with ' ' and capitalizes the form name
        if file_name[i] != '_':
            f.write(file_name[i].upper())
        else:
            f.write(' ')
    spreadsheet_sheets = excel_spreadsheet.sheet_names
    sheet = []

    for m in range(len(spreadsheet_sheets)):
        sheet.append(pd.read_excel(excel_spreadsheet, spreadsheet_sheets[m]))       #opens all the sheets in the excel file and saves them as a list of dataframes

    indices = cols[3:len(cols) - 1]  # gets a list of names of column headers in a sheet

    sheet = rename_submission__uuid(sheet)
    combined_sheet = gen_combined_sheet(sheet)
    if len(sheet)<=2:      #for excel file with only one sheet or 2 sheets
        f.write('**\n')

        for i in range(combined_sheet.shape[1]):        #this loop sorts the data in the form according to date in descending order if there is any column for date
            if combined_sheet.columns[i][0] != '_' and 'Date' in combined_sheet.columns[i]:
                combined_sheet[combined_sheet.columns[i]] = pd.to_datetime(combined_sheet[combined_sheet.columns[i]])
                combined_sheet = combined_sheet.sort_values(by=[combined_sheet.columns[i]], ascending = False)
                break

        for i in range(combined_sheet.shape[1]):        #this loop changes the format of dates to 'mm.dd.yyyy' with a separator of user's choice
            if combined_sheet.columns[i][0] != '_' and 'Date' in combined_sheet.columns[i]:
                combined_sheet[combined_sheet.columns[i]] = pd.to_datetime(combined_sheet[combined_sheet.columns[i]])
                combined_sheet[combined_sheet.columns[i]] = combined_sheet[combined_sheet.columns[i]].dt.strftime("%d"+separator+"%m"+ separator + "%Y")

        for i in range(combined_sheet.shape[0]):        #converts all the elements of the dataframe to str format
            for j in range(combined_sheet.shape[1]):
                combined_sheet.iloc[i,j] = str(combined_sheet.iloc[i,j])
        
        col_loc = column_location(combined_sheet, indices)

        for i in range(combined_sheet.shape[0]):    #this loop writes the data in report.md
            f.write('- ')
            for j in range(len(col_loc)-1):
                f.write(combined_sheet.iloc[i,col_loc[j]] + ', ')
            f.write(combined_sheet.iloc[i,col_loc[len(col_loc)-1]])
            f.write('\n')
        f.write('<br><br><br>\n')

    else:                               #for excel files having more than two sheets i.e. for forms that differ with category
        f.write(' - ' + cols[2].replace('_', ' ').upper() + '**\n')

        for j in range(len(combined_sheet)):            #this loop sorts the data in the form according to date in descending order if there is any column for date
            for i in range(combined_sheet[j].shape[1]):
                if combined_sheet[j].columns[i][0] != '_' and 'Date' in combined_sheet[j].columns[i]:
                    combined_sheet[j][combined_sheet[j].columns[i]] = pd.to_datetime(combined_sheet[j][combined_sheet[j].columns[i]])
                    combined_sheet[j] = combined_sheet[j].sort_values(by=[combined_sheet[j].columns[i]], ascending = False)
                    break

        for j in range(len(combined_sheet)):        #this loop changes the format of dates to 'mm.dd.yyyy' with a separator of user's choice
            for i in range(combined_sheet[j].shape[1]):
                if combined_sheet[j].columns[i][0] != '_' and 'Date' in combined_sheet[j].columns[i]:
                    combined_sheet[j][combined_sheet[j].columns[i]] = pd.to_datetime(combined_sheet[j][combined_sheet[j].columns[i]])
                    combined_sheet[j][combined_sheet[j].columns[i]] = combined_sheet[j][combined_sheet[j].columns[i]].dt.strftime("%d"+separator+"%m"+ separator + "%Y")

        for i in range(len(combined_sheet)):        #converts all the elements of the dataframe to str format
            for j in range(combined_sheet[i].shape[0]):
                for k in range(combined_sheet[i].shape[1]):
                    combined_sheet[i].iloc[j, k] = str(combined_sheet[i].iloc[j,k])

        for l in range(2, len(sheet)):              #this loop first matches the sheetname of the excel file to the sheetname in the template to get the correct column headers for the correct sheet  
            if cols[2] == spreadsheet_sheets[l]:
                break
        col_loc = column_location(combined_sheet[l - 2], indices)  

        for i in range(combined_sheet[l-2].shape[0]):       #writes the data in list format for each individual category
            f.write('- ')
            for j in range(len(col_loc) - 1):
                f.write(combined_sheet[l-2].iloc[i, col_loc[j]] + ', ')
            f.write(combined_sheet[l-2].iloc[i, col_loc[len(col_loc) - 1]])
            f.write('\n')
        f.write('<br><br><br>\n')



def xls2md_table(excel_spreadsheet, file_name, cols, f, separator):
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
    indices = cols[3:len(cols) - 1]               # gets a list of names of column headers in a sheet

    sheet = rename_submission__uuid(sheet)
    combined_sheet = gen_combined_sheet(sheet)

    if len(sheet)<=2:               # merges the sheets in the excel file according to the column '_uuid'
        f.write('**\n')

        for i in range(combined_sheet.shape[1]):   #sorts the data in the dataframe according to date.
            if combined_sheet.columns[i][0] != '_' and 'Date' in combined_sheet.columns[i]:
                combined_sheet[combined_sheet.columns[i]] = pd.to_datetime(combined_sheet[combined_sheet.columns[i]])
                combined_sheet = combined_sheet.sort_values(by = [combined_sheet.columns[i]], ascending = False)
                break

        for i in range(combined_sheet.shape[1]):   #converts the date format to ddmmyyyy with a separator of user's choice
            if combined_sheet.columns[i][0] != '_' and 'Date' in combined_sheet.columns[i]:
                combined_sheet[combined_sheet.columns[i]] = pd.to_datetime(combined_sheet[combined_sheet.columns[i]])
                combined_sheet[combined_sheet.columns[i]] = combined_sheet[combined_sheet.columns[i]].dt.strftime("%d"+separator+"%m"+ separator + "%Y")

        for i in range(combined_sheet.shape[0]):    #converts all data of the dataframe to string format
            for j in range(combined_sheet.shape[1]):
                combined_sheet.iloc[i,j] = str(combined_sheet.iloc[i,j])

        for i in range(len(indices)):   #writes the column headers in the report.md file
            f.write('| ' + indices[i])
        f.write(' |\n')

        for i in range(len(indices)):   #writes the separator for column headers and data in report.md file
            f.write('|-----')
        f.write(' |\n')
        
        col_loc = column_location(combined_sheet, indices)

        for i in range(combined_sheet.shape[0]):    #writes the data in report.md file
            for j in range(len(col_loc)):
                for k in range(len(combined_sheet.iloc[i, col_loc[j]])):
                    if combined_sheet.iloc[i, col_loc[j]][k] == '\n':   #checks if there are any lists in data and separates them accordingly
                        combined_sheet.iloc[i, col_loc[j]] = combined_sheet.iloc[i, col_loc[j]][:k-1] + ' <br> ' + combined_sheet.iloc[i, col_loc[j]][k+1:]
                f.write('| ' + combined_sheet.iloc[i,col_loc[j]])
            f.write(' |\n')
        f.write('<br><br><br>\n')
    else:
        f.write(' - ' + cols[2].replace('_', ' ').upper() + '**\n') #writes sheet naem along with form name for multi-category forms

        for j in range(len(combined_sheet)):    #sorts the data according to date in descending order
            for i in range(combined_sheet[j].shape[1]):
                if combined_sheet[j].columns[i][0] != '_' and 'Date' in combined_sheet[j].columns[i]:
                    combined_sheet[j][combined_sheet[j].columns[i]] = pd.to_datetime(combined_sheet[j][combined_sheet[j].columns[i]])
                    combined_sheet[j] = combined_sheet[j].sort_values(by=[combined_sheet[j].columns[i]], ascending = False)
                    break

        for j in range(len(combined_sheet)):        #changes th fprmat of date to ddmmyyyy with a separator of user's choice
            for i in range(combined_sheet[j].shape[1]):
                if combined_sheet[j].columns[i][0] != '_' and 'Date' in combined_sheet[j].columns[i]:
                    combined_sheet[j][combined_sheet[j].columns[i]] = pd.to_datetime(combined_sheet[j][combined_sheet[j].columns[i]])
                    combined_sheet[j][combined_sheet[j].columns[i]] = combined_sheet[j][combined_sheet[j].columns[i]].dt.strftime("%d"+separator+"%m"+ separator + "%Y")

        for i in range(len(combined_sheet)):        #converts all data pf dataframe to string format
            for j in range(combined_sheet[i].shape[0]):
                for k in range(combined_sheet[i].shape[1]):
                    combined_sheet[i].iloc[j, k] = str(combined_sheet[i].iloc[j, k])

        for i in range(len(indices)):       #writes column headers in report.md
            f.write('| ' + indices[i])
        f.write(' |\n')

        for i in range(len(indices)):       #writes separatoe for column headers and data in report.md
            f.write('|-----')
        f.write(' |\n')

        for l in range(2, len(sheet)):      #checks for the correct sheet from template.txt file
            if cols[2] == spreadsheet_sheets[l]:
                break
        col_loc = column_location(combined_sheet[l-2], indices)

        for i in range(combined_sheet[l-2].shape[0]):   #writes the data in report.md file
            for j in range(len(col_loc)):
                for k in range(len(combined_sheet[l-2].iloc[i, col_loc[j]])):
                    if combined_sheet[l-2].iloc[i, col_loc[j]][k] == '\n':      #cheks if there are any lists in the spreadhseet and separates them accordingly
                        combined_sheet[l-2].iloc[i, col_loc[j]] = combined_sheet[l-2].iloc[i, col_loc[j]][:k - 1] + ' <br> ' + combined_sheet[l-2].iloc[i, col_loc[j]][k + 1:]
                f.write('| ' + combined_sheet[l-2].iloc[i, col_loc[j]])
            f.write(' |\n')
        f.write('<br><br><br>\n')


cur_path = os.getcwd()  # gets the path where the python file is located
template_path = cur_path + r'\template'
path = cur_path + r'\excel'
template = open(template_path + '\\' + 'template.txt')
template_fields = template.readlines()  #generates a list of all templates in the template.txt file

for i in range(len(template_fields)):   #separates individual form templates so that they can be easily used
    template_fields[i] = template_fields[i].split(',')

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
    if template_fields[k][0] == 't':                                                                                                    
        xls2md_table(spreadsheet, filename[i], template_fields[k],f, sep)     
    elif template_fields[k][0] == 'l':
        xls2md_list(spreadsheet, filename[i], template_fields[k],f, sep)

print("All Files have been converted to Markdown")
