import xlwings as xw
import pandas as pd


def load_data_table_to_df(xl_file_name, sheet_name, table_name):
    """
    Load an Excel data table content to a pandas dataframe
    :param xl_file_name: Excel file path & name
    :param sheet_name: sheet name where the data table(s) can be found
    :param table_name: table name
    :return: pandas dataframe or in case of error None
    """
    # Start invisible Excel
    xl_app = xw.App(visible=False, add_book=False)
    wb = None

    try:
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

    except Exception as ex:
        template = "An exception of type {0} occurred. Arguments:\n{1!r}"
        message = template.format(type(ex).__name__, ex.args)
        print(message)

        if wb is not None:
            wb.close()

        xl_app.quit()


if __name__ == '__main__':

    df_test1 = load_data_table_to_df('data_tbl_test.xlsx', 'test_datatbl_1', 'tbl_test_3')
    df_test2 = load_data_table_to_df('data_tbl_test.xlsx', 'test_datatbl_2', 'tbl_test_1')
    df_test3 = load_data_table_to_df('data_tbl_test.xlsx', 'test_datatbl_2', 'tbl_test_2')

    if df_test1 is not None:
        print(df_test1.head())

    if df_test2 is not None:
        print(df_test2.head())

    if df_test3 is not None:
        print(df_test3.head())
