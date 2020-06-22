"""
Report generator

This script reads the template from 'template.txt' file in the 'template' folder and generates the report for all the files present in 'excel folder
according to the template provided by the user. The report is saved in a file 'report.md' in markdown folder. The report
is generated in Markdown format.

This script works only with '.xlsx' files.

For this script to work, 'pandas' and 'glob' must be installed.

This script contains the following functions:
   * column_location - returns a list conatining indices of columns in correct sequence according to 'template.txt' file.
   * date_sort_and_format - returns list of dataframes of all sheets of an excel file after sorting the data in
                            descending order according to date column (if any) in a sheet. It also converts the date
                            format to ddmmyyyy with a separator of yser's choice.
   * gen_list_sheets - returns a lost of dataframes after combining relevant sheets of the excel file using the '_uuid'
                       column which contains the unique id for each submission.
   * write_table - writes a dataframe in tabular format in report.md.
   * write_list - writes a dataframe in list format in report.md.
   * xls2md - this function determines whether the data for a particular excel file is to be writeen in tabular or list
              format and invokes wrute_table or write_list functions accordingly.
   * rename_submission__uuid - this function renames '_submission__uuid' column to '_uuid' and returns the list of
                               dataframes with renamed column.
"""


import pandas as pd
import os
import glob
import datetime


def column_location(spreadsheet, column_names):
    """
    This function finds the indices of columns in the order specified in template file.

        Parameters:
            spreadsheet (DataFrame): A pandas DataFrame.
            column_names (List of srtings): list of strings that specifies the name of column headers as in template file.

        Returns:
            column_location (list of integers): list of integers specifying the column locations in the DataFrame according tp the template specified.
    """
    column_location = []
    for i in range(len(column_names)):
        for j in range(spreadsheet.shape[1]):
            if column_names[i] in spreadsheet.columns[j]:
                column_location.append(j)
    return column_location


def date_sort_and_format(list_sheets, separator):
    """
    Returns the DataFrame after sorting the data in descending order of date or year column if there is any column that
    contains date or simply year in the DataFrame.

        Parameters:
            list_sheets (list of dataframes): a list of Pandas DataFrames.
            separator (character): a special character to separate the day month and year

        Returns:
            all_sheets (list of DataFrames): a list of DataFrames with all data sorted in descending order of date column(if any) and format of date converted to ddmmyyyy
    """
    all_sheets = []
    for i in range(len(list_sheets)):
        all_sheets.append(list_sheets[i])

    for j in range(len(all_sheets)):
        for i in range(all_sheets[j].shape[1]):
            if all_sheets[j].columns[i][0] != '_' and 'Date' in all_sheets[j].columns[i]: # finds the column conataining date
                all_sheets[j][all_sheets[j].columns[i]] = pd.to_datetime(all_sheets[j][all_sheets[j].columns[i]])
                all_sheets[j] = all_sheets[j].sort_values(by=[all_sheets[j].columns[i]], ascending=False)
                break
            elif all_sheets[j].columns[i][0] != '_' and 'Year' in all_sheets[j].columns[i]: # finds the column containing year
                all_sheets[j] = all_sheets[j].sort_values(by=[all_sheets[j].columns[i]], ascending=False)
                break

    for j in range(len(all_sheets)):
        for i in range(all_sheets[j].shape[1]):
            if all_sheets[j].columns[i][0] != '_' and 'Date' in all_sheets[j].columns[i]:
                all_sheets[j][all_sheets[j].columns[i]] = pd.to_datetime(all_sheets[j][all_sheets[j].columns[i]])
                all_sheets[j][all_sheets[j].columns[i]] = all_sheets[j][all_sheets[j].columns[i]].dt.strftime(
                    "%d" + separator + "%m" + separator + "%Y")
    return all_sheets


def gen_list_sheets(sheets):
    """
    Returns a list of DataFrames which can be worked upon to generate reports.

    This function takes a list of dataframes which are all the sheets in an excel file and merges the relevant sheets
    according to the unique id column '_uuid'

        Parameters:
            sheets (list of DataFrames): A list of pandas DataFrames. All the sheets of the excel file opened as DataFrames and stored in a list

        Returns:
            all_sheets (List of DataFrames):A list of DataFrames after performing the merging operation.
    """
    all_sheets = []
    if len(sheets) == 1:       # if there is only one sheet in the file then it returns the sheet itself
        all_sheets.append(sheets[0])
        return all_sheets  # returns the first sheet if there is only one
    all_sheets.append(pd.merge(sheets[0], sheets[1], on='_uuid'))
    if (len(sheets)) == 2:     # if the form has 2 sheets then both are merged according to '_uuid' column and #
        return all_sheets      # returned as a dataframe
    for j in range(len(sheets) - 2):
        all_sheets.append(pd.merge(all_sheets[0], sheets[j + 2], on='_uuid'))
                                    # if more than 2 sheets i.e. multi-category form then combines 3rd sheet onwards
    return all_sheets[1:]           # to 1st and 2nd sheet and returns a list of combined sheets


def write_table(combined_sheet, column_names, f):
    """
    Writes the data of DataFrame in tabular format in Markdown.

        Parameters:
            combined_sheet (DataFrame): A Pandas DataFrame whose data is to be written in Tabular Form in Markdown.
            column_names (List of Strings): A list of strings containing all the relevant column names of the DataFrame.
            f (file object): file object to write the data in 'report.md'
    """
    for i in range(len(column_names)):
        f.write('| ' + column_names[i])
    f.write(' |\n')

    for i in range(len(column_names)):
        f.write('|-----')
    f.write(' |\n')
    col_loc = column_location(combined_sheet, column_names)

    for i in range(combined_sheet.shape[0]):
        for j in range(len(col_loc)):
            for k in range(len(combined_sheet.iloc[i, col_loc[j]])):
                if combined_sheet.iloc[i, col_loc[j]][k] == '\n':
                    combined_sheet.iloc[i, col_loc[j]] = combined_sheet.iloc[i, col_loc[j]][:k - 1] + ' <br> ' + combined_sheet.iloc[i, col_loc[j]][k + 1:]
            f.write('| ' + combined_sheet.iloc[i, col_loc[j]])
        f.write(' |\n')
    f.write('<br><br><br>\n')


def write_list(combined_sheet, column_names, f):
    """
    Writes the data of DataFrame in list format in Markdown. The function works for basic lists. However, it currently doesn't
    work for nested list.

        Parameters:
            combined_sheet (DataFrame): A Pandas DataFrame whose data is to be written in list Form in Markdown.
            column_names (List of strings): A list of strings containing all the relevant column names of the DataFrame.
            f (file object): file object to write the data in 'report.md'
    """
    col_loc = column_location(combined_sheet, column_names)

    for i in range(combined_sheet.shape[0]):
        f.write('- ')
        for j in range(len(col_loc) - 1):
            if combined_sheet.iloc[i, col_loc[j]] == 'nan':
                continue
            else:
                f.write(combined_sheet.iloc[i, col_loc[j]] + ', ')
        if not combined_sheet.iloc[i, col_loc[len(col_loc) - 1]] == 'nan':
            f.write(combined_sheet.iloc[i, col_loc[len(col_loc) - 1]])
        f.write('\n')
    f.write('<br><br><br>\n')


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


def xls2md(excel_spreadsheet, file_name, template, f, separator):
    """
    Writes all data to 'reprt.md' file in markdown format

    This function takes in the excel file, its name, its specified template, file object and separator to write in dates
    and wrtes that data in tabular or list format as specified in the template.

        Parameters:
            excel_spreadsheet (DataFrame): DataFrame of all the excel spreadsheeets in a file.
            file_name (string): String specifying name of the excel file.
            template (list of strings): A list of string specifying the template for the Excel file
            f (file object): File object to write data in Markdown format in the file 'report.md'
            separator (character): A character for separating the day, month and year for the dates in the report.

    """

    sheetnames = excel_spreadsheet.sheet_names
    sheets = []

    for m in range(len(sheetnames)):  # stores all sheets in an excel file in a list 'sheet'
        sheets.append(pd.read_excel(excel_spreadsheet, sheetnames[m]))
    column_names = template[3:len(template) - 1]  # gets a list of names of column headers in a sheet

    sheets = rename_submission__uuid(sheets)
    list_sheets = gen_list_sheets(sheets)
    list_sheets = date_sort_and_format(list_sheets, separator)

    f.write('## **')

    for i in range(file_name.index('_-_')):  # determines the form name from file name and converts it to all caps
        if file_name[i] != '_':  # and writes it in 'report.md'
            f.write(file_name[i].upper())
        else:
            f.write(' ')

    if len(sheets)<=2:         # if the form has single set of repeating questions or no repeating questions
        f.write('**\n')
        for i in range(list_sheets[0].shape[0]):
            for j in range(list_sheets[0].shape[1]):
                list_sheets[0].iloc[i, j] = str(list_sheets[0].iloc[i, j])
        if template[0] == 't':
            write_table(list_sheets[0], column_names, f)
        elif template[0] == 'l':
            write_list(list_sheets[0], column_names, f)
    else:             # if the form has multiple sets of repeating questions the=at vary according to category selected.
        f.write(' - ' + template[2].replace('_', ' ').upper() + '**\n')
        for l in range(2, len(sheets)):     # checks for the correct sheetname as in the template
            if template[2] == sheetnames[l]:
                break
        for i in range(list_sheets[l-2].shape[0]):
            for j in range(list_sheets[l-2].shape[1]):
                list_sheets[l-2].iloc[i, j] = str(list_sheets[l-2].iloc[i, j])
        if template[0] == 't':          # l-2 because the categories of form start from sheet number 3.
            write_table(list_sheets[l - 2], column_names, f)
        elif template[0] == 'l':
            write_list(list_sheets[l - 2], column_names, f)


def main():
    cur_path = os.getcwd()
    template_path = cur_path + r'\template'
    path = cur_path + r'\excel'
    template = open(template_path + '\\' + 'template.txt')
    template_fields = template.readlines()      # gets the template for all the forms from 'template.txt' file line by line

    for i in range(len(template_fields)):
        template_fields[i] = template_fields[i].split(';')

    files = [f for f in glob.glob(path + '**/*.xlsx',recursive=True)]  # gets the paths for all the excel files stored in 'excel' folder
    filename = []
    lenpath = len(path) + 1
    dest_path = os.getcwd() + '\\'
    dest_path = dest_path + 'markdown\\'
    dest_path = dest_path + 'report.md'  # destination file for markdown
    sep = str(input("which separator do you want for dates? Please Enter\n"))
    f = open(dest_path, 'w')

    for file in files:
        filename.append(file[lenpath:len(file) - 5])  # gets the filenames of all the files in the 'excel' folder

    for k in range(len(template_fields)):
        for i in range(len(filename)):
            if filename[i] in template_fields[k][1]:
                break
        spreadsheet = pd.ExcelFile(files[i])  # opens all the excel files one by one
        xls2md(spreadsheet, filename[i], template_fields[k], f, sep)  # the function that converts the excel spreadsheet to md

    print("All Files have been converted to Markdown")


if __name__ == '__main__':
    main()
