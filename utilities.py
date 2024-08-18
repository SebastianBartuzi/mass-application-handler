import openpyxl

EXCEL_FILE_PATH = 'debug/input/data.xlsx'


def replace_townname(towns_data, i):
    workbook = openpyxl.load_workbook(EXCEL_FILE_PATH)
    sheet = workbook["Dane"]
    town_name = towns_data[i][0]

    if towns_data[i][2].upper() == "W" or towns_data[i][2].upper() == "B":
        town_name = "Gmina " + town_name
    elif towns_data[i][2].upper() == "P":
        town_name = "Miasto " + town_name
    sheet[f'B{i + 16}'] = town_name
    workbook.save(EXCEL_FILE_PATH)
    workbook.close()
    return town_name


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
        if town_data[0] is None:  # Stop if the cell is empty
            break
        towns_data.append(town_data)
        row += 1

    return user_data, towns_data
