import openpyxl

EXCEL_FILE_PATH = 'debug/input/data.xlsx'

def read_excel():
    # Load the Excel workbook
    workbook = openpyxl.load_workbook(EXCEL_FILE_PATH)

    # Select the active worksheet (you can also specify a worksheet by name)
    sheet = workbook["Dane"]

    # Initialize an empty list to store the cell values
    user_data = []
    user_data.append(sheet['F4'].value)
    user_data.append(sheet['G4'].value)
    user_data.append(sheet['H4'].value)
    user_data.append(sheet['I4'].value)
    user_data.append(sheet['J4'].value)

    # Start creating files from row 16
    row = 16
    towns_data = []
    while True:
        town_data = []
        town_data.append(sheet[f'B{row}'].value)
        town_data.append(sheet[f'C{row}'].value)
        town_data.append(sheet[f'D{row}'].value)
        town_data.append(sheet[f'E{row}'].value)
        town_data.append(sheet[f'F{row}'].value)
        town_data.append(sheet[f'G{row}'].value)
        town_data.append(sheet[f'H{row}'].value)
        if town_data[0] is None:  # Stop if the cell is empty
            break
        towns_data.append(town_data)
        row += 1

    return user_data, towns_data
