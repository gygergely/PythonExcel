import xlwings as xw
import pandas as pd


def load_data_table_to_df(xl_file_name, sheet_name, table_name):
    """

    :param xl_file_name:
    :param sheet_name:
    :param table_name:
    :return:
    """
    # Start invisible Excel
    xl_app = xw.App(visible=False, add_book=False)

    # Open source file
    wb = xl_app.books.open(xl_file_name)

    # Identify the sheet
    sh = wb.sheets[sheet_name]

    # Get the table by name
    data_tbl = sh.api.ListObjects(table_name)

    # Get the table range
    table_range = sh.range(data_tbl.range.address)

    # Load the table range values to a dataframe
    df = pd.DataFrame(table_range.value)

    # Grab the first row for the header
    df_header = df.iloc[0]

    # Get the data except the 1st row
    df = df[1:]

    # Set the 1st row as header
    df.columns = df_header

    # Reset df index
    df.reset_index(drop=True, inplace=True)

    # Close Excel
    wb.close()
    xl_app.quit()

    return df


if __name__ == '__main__':

    df_test1 = load_data_table_to_df('data_table_test.xlsx', 'test_datatbl_1', 'tbl_test_4')
    df_test2 = load_data_table_to_df('data_table_test.xlsx', 'test_datatbl_2', 'tbl_test_1')
    df_test3 = load_data_table_to_df('data_table_test.xlsx', 'test_datatbl_2', 'tbl_test_2')

    print(df_test1.head())
    print(df_test2.head())
    print(df_test3.head())
