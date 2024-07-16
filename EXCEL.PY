import openpyxl
workbook = openpyxl.load_workbook('FIILE_NAME.xlsx')
sheet = workbook.active
for row in sheet.iter_rows():
    for cell in row:
        if cell.data_type == 'e':
            print(f"Error in cell {cell.coordinate}: {cell.value}")