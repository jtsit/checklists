#this is the concept in python.

import csv  # to read csv file
import xlsxwriter  # to write xlxs file
import ast

# you can change this names according to your local ones
csv_file = 'data.csv'
xlsx_file = 'data.xlsx'

# read the csv file and get all the JSON values into data list
data = []
with open(csv_file, 'r') as csvFile:
    # read line by line in csv file
    reader = csv.reader(csvFile)

    # convert every line into list and select the JSON values
    for row in list(reader)[1:]:
        # csv are comma separated, so combine all the necessary
        # part of the json with comma
        json_to_str = ','.join(row[1:])

        # convert it to python dictionary
        str_to_dict = ast.literal_eval(json_to_str)

        # append those completed JSON into the data list
        data.append(str_to_dict)

# define the excel file
workbook = xlsxwriter.Workbook(xlsx_file)

# create a sheet for our work
worksheet = workbook.add_worksheet()

# cell format for merge fields with bold and align center
# letters and design border
merge_format = workbook.add_format({
    'bold': 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter'})

# other cell format to design the border
cell_format = workbook.add_format({
    'border': 1,
})

# create the header section dynamically
first_col = 0
last_col = 0
for index, value in enumerate(data[0].items()):
    if isinstance(value[1], dict):
        # this if mean the JSON key has something else
        # other than the single value like dict or list
        last_col += len(value[1].keys())
        worksheet.merge_range(first_row=0,
                              first_col=first_col,
                              last_row=0,
                              last_col=last_col,
                              data=value[0],
                              cell_format=merge_format)
        for k, v in value[1].items():
            # this is for go in deep the value if exist
            worksheet.write(1, first_col, k, merge_format)
            first_col += 1
        first_col = last_col + 1
    else:
        # 'age' has only one value, so this else section
        # is for create normal headers like 'age'
        worksheet.write(1, first_col, value[0], merge_format)
        first_col += 1

# now we know how many columns exist in the
# excel, and set the width to 20
worksheet.set_column(first_col=0, last_col=last_col, width=20)

# filling values to excel file
for index, value in enumerate(data):
    last_col = 0
    for k, v in value.items():
        if isinstance(v, dict):
            # this is for handle values with dictionary
            for k1, v1 in v.items():
                if isinstance(v1, list):
                    # this will capture last 'type' list (['Grass', 'Hardball'])
                    # in the 'conditions'
                    worksheet.write(index + 2, last_col, ', '.join(v1), cell_format)
                else:
                    # just filling other values other than list
                    worksheet.write(index + 2, last_col, v1, cell_format)
                last_col += 1
        else:
            # this is handle single value other than dict or list
            worksheet.write(index + 2, last_col, v, cell_format)
            last_col += 1

# finally close to create the excel file
workbook.close()
