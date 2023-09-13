import openpyxl


def find_excel_cell(workbook_path, worksheet_name, search_text, case_sensitive=False):
    
    # load workbook and worksheet
    try:
        workbook = openpyxl.load_workbook(workbook_path, data_only=True)
        worksheet = workbook[worksheet_name]
    except Exception as e:
        print(e)
        return

    coordinates = []

    # search through the sheet
    for row in worksheet.iter_rows():
        for cell in row:
            cell_value = cell.value

            if not case_sensitive:
                cell_value = str(cell_value).upper()
                search_text = search_text.upper()

            # check if the cell value matches the search text
            if cell_value == search_text:

                # get indices
                row_index = cell.row - 1
                column_index = cell.column - 1
                coordinates.append((row_index, column_index))

    if len(coordinates) != 0:
        print(f"The coordinates of '{search_text}' are {coordinates}")
    else:
        print(f"'{search_text}' not found in the worksheet.")


# example:
workbook_path = "TothAdam_ExcelFeladat.xlsx"
worksheet_name = "1. feladat"
search_text = "nagyobb"
case_sensitive = True

find_excel_cell(workbook_path, worksheet_name, search_text, case_sensitive=case_sensitive)

