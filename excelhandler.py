__author__ = 'Om Computers'

import xlrd
from xlrd import cellname

def open_xl(file_name, sheet_name):
    data = xlrd.open_workbook(file_name)
    sheet = data.sheet_by_name(sheet_name)
    return sheet

def print_xl_allVals(sheet):
    for row in range(sheet.nrows):
        for col in range(sheet.ncols):
            # cell_name = cellname(row, col)
            print sheet.cell(row, col).value

def print_xl_specificCol_in_allRows(sheet, specific_col_index):
    name_salaries_dict = {}
    salaries = []
    for row in range(1, sheet.nrows):
        salary = sheet.cell(row, specific_col_index).value
        name = sheet.cell(row, 0).value
        name_salaries_dict[name] = salary
        salaries.append(salary)
    return name_salaries_dict, salaries

def highest_earning(name_salaries_dict, salaries):
    max_salary = max(salaries)
    final_str = ''
    for name, sal in name_salaries_dict.iteritems():
        if sal == max_salary:
            final_str = final_str + name + " earns " + str(sal) + " USD." + '\n'
    return final_str

def write_to_file(final_str):
    f = open('out.txt', 'w')
    f.write(final_str)
    f.close()


#xlfile = "Financial Sample.xlsx"
xlfile = "fin.xlsx"
sheet_name = 'Sheet1'
sheet = open_xl(xlfile, sheet_name)
#print_xl_allVals(sheet)
name, sal = print_xl_specificCol_in_allRows(sheet, 2)
final_str = highest_earning(name, sal)
write_to_file(final_str)