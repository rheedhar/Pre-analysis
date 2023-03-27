from functions import *

# fetch all documents
workbook, main_sheet, data_wb, data_ws, final_ws = create_output_sheet()

# check if the number of columns is the same in both the result file and the data file.
columns_inserted = False
for row1 in main_sheet.iter_rows(min_row=1, min_col=2):
    for row2 in data_ws.iter_rows(min_row=1, min_col=2):
        if len(row1) != len(row2):
            columns_inserted = insert_column(min_row=1, min_col=2, start_insert=3, result_sheet=main_sheet)
            break

# unmerge any merged cells in the result sheet and put new values to the new cells if they are empty
merged_cells = handle_merged_cells(result_sheet=main_sheet)

# paste data into result shells
j = 0
last_column_coord = ""
for row in main_sheet.iter_rows(min_row=1, min_col=1):
    row_values = tuple(row_cell.value for row_cell in row)
    if any(map(lambda x: x is None, row_values)):
        continue

    # #grab the last coordinate with data
    last_column_coord = row[-1].coordinate
    for index2, row2 in enumerate(data_ws.iter_rows(min_row=2, min_col=2)):
        if index2 == j:
            for index, cell in enumerate(row[1:]):
                main_sheet[cell.coordinate] = row2[index].value
            j += 1
            break

# sort columns that are highlighted.
# first we need to fetch the range of columns that are highlighted, if there are highlighted columns,
# and then sort the columns afterwards.
for col in main_sheet.iter_cols(min_row=1, min_col=1, max_col=1):
    if col.fill.start_color.index == "FFFFFF00":
        highlighted_columns = fetch_highlighted_columns(main_sheet, last_column_coord)
        sort_rows_data(main_sheet, highlighted_columns)
        break

# format final result sheet.
for row in main_sheet.iter_rows(min_row=1, min_col=1):
    row_values = tuple(row_cell.value for row_cell in row)
    if any(map(lambda x: x is None, row_values)):
        continue

    label_value = row[0].coordinate.value
    grouped_cells = combine_cells(row[1:])

    for cell in grouped_cells:
        first_cell, second_cell, first_value, second_value = split_cells(cell)
        if merged_cells:
            if "mean" in label_value.lower():
                main_sheet.merge_cells(cell)
                main_sheet[cell.split(":")[0]] = f"{first_value}({second_value})"
            elif any(value in label_value.lower() for value in ("95% ci", "q1, q3", "minimum, maximum", "min, max")):
                main_sheet.merge_cells(cell)
                main_sheet[cell.split(":")[0]] = f"({first_value}, {second_value})"
            elif any(
                    value in label_value.lower() for value in ("median", "total patient sample", "number of patients")):
                main_sheet.merge_cells(cell)
                main_sheet[cell.split(":")[0]] = f"{first_value}"
            elif any(value in label_value.lower() for value in ("12", "24", "36", "48", "60")):
                main_sheet[cell.split(":")[0]] = f"{first_value}%"
                main_sheet[cell.split(":")[1]] = f"{second_value}"
            else:
                main_sheet[cell.split(":")[0]] = f"{first_value}"
                main_sheet[cell.split(":")[1]] = f"{second_value}%"
        else:
            if "mean" in label_value.lower():
                main_sheet[cell.split(":")[0]] = f"{first_value}({second_value})"
            elif any(value in label_value.lower() for value in ("95% ci", "q1, q3", "minimum, maximum", "min, max")):
                main_sheet[cell.split(":")[0]] = f"({first_value}, {second_value})"
            elif any(
                    value in label_value.lower() for value in ("median", "total patient sample", "number of patients")):
                main_sheet[cell.split(":")[0]] = f"{first_value}"
            elif any(value in label_value.lower() for value in ("12", "24", "36", "48", "60")):
                main_sheet[cell.split(":")[0]] = f"{first_value}%({second_value})"
            else:
                main_sheet[cell.split(":")[0]] = f"{first_value}({second_value}%)"

# delete columns that were inserted previously
if columns_inserted:
    delete_column(main_sheet)

# save results
workbook.save(f"data/{final_ws}.xlsx")