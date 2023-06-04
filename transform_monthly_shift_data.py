import openpyxl
from generate_sample_shiftdata import non_empty_file_path, output_workbook, output_file_path

# Load the workbook
worksheet_path = input("Enter the path to the input Excel file: ")
worksheet_path = non_empty_file_path(worksheet_path, "Enter the path to the input Excel file: ")
workbook = openpyxl.load_workbook(worksheet_path)

# Select the desired sheet
source_sheet = workbook['Sheet1']
new_sheet = output_workbook['Sheet']

legends = {
    'M': '07:00-16:00',
    'G': '09:00-18:00',
    'A': '14:00-23:00',
    'L': 'Leave',
    'H': 'Holiday',
    'WO': 'Weekly Off'
}


# Iterate over the cells in the source sheet
def get_shift_value(cell_val):
    if cell_val == 'M':
        return legends[cell_val]
    elif cell_val == 'G':
        return legends[cell_val]
    elif cell_val == 'A':
        return legends[cell_val]
    elif cell_val == 'L':
        return legends[cell_val]
    elif cell_val == 'H':
        return legends[cell_val]
    elif cell_val == 'WO':
        return legends[cell_val]
    else:
        return cell_val


for row in source_sheet.iter_rows():
    for cell in row:
        value = get_shift_value(cell.value)
        new_cell = new_sheet.cell(row=cell.row, column=cell.column, value=value)



# Save the workbook
output_workbook.save(output_file_path)
