import openpyxl

def get_values_range_from_excel(file_path, sheet_name, start_cell, end_cell):
    workbook = openpyxl.load_workbook(file_path, data_only=True)
    sheet = workbook[sheet_name]

    sum_str = ""
    for row in sheet[start_cell:end_cell]:
        for cell in row:
            sum_str += cell.value
    return sum_str

file_path = r'C:\Users\TechnoStar\Documents\macro\funnction\New Microsoft Excel Worksheet.xlsx'

cell_values = get_values_range_from_excel(file_path, 'Sheet1', 'A1', 'A10')

print(cell_values)
