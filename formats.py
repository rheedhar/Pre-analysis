from format_functions import extract_sheets, create_data_struct, create_format_file


def main():
    # fetch documents
    workbook, worksheet, formats_file = extract_sheets()

    # create data dictionary
    format_data = create_data_struct(worksheet)

    # create format file
    create_format_file(format_data, formats_file)


if __name__ == "__main__":
    main()