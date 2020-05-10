import xlwings as xw


def read_named_range_from_sheet(xl_file_name, sheet_name, list_of_named_ranges, period_id, reporting_id):
    """
    read non contiguous ranges from an excel file to a list
    :param xl_file_name: the Excel file name
    :param sheet_name: sheet name holding the named ranges
    :param list_of_named_ranges: list of the named ranges to read, 1st name must be the 'row' range, the others are
    the value ranges example:
    Excel:
        name    Q1-value    Q2-value    Q3-value    Q4-value
    Expected output:
        name    Q1-value
        name    Q2-value
        name    Q3-value
        name    Q4-value
    :param period_id: period_ids to be assigned to the value ranges
    :param reporting_id: reporting period identification
    :return: list in a format of expected output
    """
    # Open template file
    wb = xl_app.books.open(xl_file_name)
    # Idnetify the sheet
    # TODO: handle if a sheet is not existing
    sh = wb.sheets[sheet_name]

    # Create empty list for results
    final_list = list()

    # Iterate through the named ranges, except the 1st one which is the 'row' range
    for idx in range(1, len(list_of_named_ranges)):

        # At each iteration create the row range and delete None values
        # TODO: handle if a named range is not existing
        rng = sh.range(list_of_named_ranges[0])
        main_row_list = rng.value
        main_row_list = [item for item in main_row_list if item[0] is not None]

        # Store the value ranges in a list
        rng = sh.range(list_of_named_ranges[idx])
        value_list = rng.value

        # Add the values to the 'row' range
        for idx_value in range(0, len(main_row_list)):
            if value_list[idx_value] is None:
                main_row_list[idx_value].append(0)
            else:
                main_row_list[idx_value].append(value_list[idx_value])

            # Add period and reporting IDs to the 'row' range
            main_row_list[idx_value].extend([period_id[idx-1], reporting_id])

        # Extend the result list with the new values
        final_list.extend(main_row_list.copy())
        main_row_list.clear()

    wb.close()
    xl_app.quit()
    return final_list


if __name__ == '__main__':
    # Start Visible Excel
    xl_app = xw.App(visible=True, add_book=False)
    named_ranges = ['name_year', 'rng_q1', 'rng_q2', 'rng_q3', 'rng_q4']
    period_ids = ['Q1', 'Q2', 'Q3', 'Q4']
    # TODO: loop through files in a folder
    result_list = read_named_range_from_sheet('test_read_ing_xlsb.xlsb', 'Test', named_ranges, period_ids, 'Q1')
    # TODO: loading to sqlite
    print(result_list)
