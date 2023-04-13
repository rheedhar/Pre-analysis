import main_functions


def main():
    # fetch all documents
    workbook, shell_sheet, data_workbook, data_sheet, output_sheet = main_functions.extract_sheets()

    # check if the number of columns is the same in both the result file and the data file.
    equal_columns = main_functions.check_equal_columns(shell_sheet, data_sheet)

    # if there aren't equal columns, insert columns to match both the data file and shell file
    if not equal_columns:
        columns_inserted = main_functions.insert_column(min_row=1, min_col=2, start_insert=3, result_sheet=shell_sheet)
    else:
        columns_inserted = False

    # check if there are merged cells in the worksheet
    merged_cells = shell_sheet.merged_cells.ranges

    # unmerge any merged cells in the result sheet and put new values to the new cells if they are empty
    if len(merged_cells) > 0:
        main_functions.handle_merged_cells(result_sheet=shell_sheet)

    # paste data into result shells and save the last column coordinate
    start_row = 0
    for shell_row in shell_sheet.iter_rows(min_row=2, min_col=1, max_row=2):
        shell_row_values = tuple(shell_cell.value for shell_cell in shell_row)
        if any(map(lambda x: x is None, shell_row_values)):
            start_row = 1
        else:
            start_row = 3
        last_column_coord = main_functions.paste_data(shell_sheet, data_sheet, start_row)

    # if there is any  highlighted column in the sheet, check for all the highlighted columns
    # and then sort the columns afterwards.
    highlighted_columns = main_functions.fetch_highlighted_columns(shell_sheet, last_column_coord)
    if len(highlighted_columns) > 0:
        main_functions.sort_rows_data(shell_sheet, highlighted_columns)

    # format final result sheet.
    main_functions.format_result_sheet(shell_sheet, columns_inserted, start_row)

    # delete columns that were inserted previously
    if columns_inserted:
        main_functions.delete_column(shell_sheet)

    # save results
    workbook.save(f"data/{output_sheet}")


if __name__ == "__main__":
    main()