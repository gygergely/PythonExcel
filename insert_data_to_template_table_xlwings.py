import csv
import xlwings as xw
import os
import datetime as dt
import shutil


def open_csv_file(csv_file_path):
    """
    Open and read data from a csv file without headers (skipping the first row)
    :param csv_file_path: path of the csv file to process
    :return: a list with the csv content
    """
    with open(csv_file_path, 'r', encoding='utf-8') as csv_file:
        reader = csv.reader(csv_file)

        # Skip header row
        next(reader)

        # Add csv content to a list
        data = list()
        for row in reader:
            data.append(row)

        return data


def write_list_to_excel(template_file, data_to_insert):
    """
    Inserting data to an existing Excel data table
    :param template_file: path of the Excel template file
    :param data_to_insert: data to insert (list)
    :return: None
    """

    try:
        # Start Visible Excel
        xl_app = xw.App(visible=True, add_book=False)

        # Open template file
        wb = xl_app.books.open(template_file)

        # Assign the sheet holding the template table to a variable
        ws = wb.sheets('TemplateTab')

        # First cell of the template (blank) table
        databody_range_first_row = 5
        databody_range_first_column = 3

        # Insert data
        ws.range((databody_range_first_row, databody_range_first_column)).value = data_to_insert

        # Save and Close the Excel template file
        wb.save()
        wb.close()

        # Close Excel
        xl_app.quit()

    except Exception as ex:
        template = "An exception of type {0} occurred. Arguments:\n{1!r}"
        message = template.format(type(ex).__name__, ex.args)
        print(message)


def check_output_folder():
    """
    Checks if 'Output' folder exists, if not it creates one inside your project
    :return: None
    """
    if not os.path.exists('Output'):
        os.makedirs('Output')


def create_date_id():
    """
    Creates a str dateID from today's date
    :return: "YYYYMMDD' format str dateID
    """
    today_date = dt.datetime.today()
    year = str(today_date.year)

    month = str(today_date.month)
    if len(month) == 1:
        month = '0' + month

    day = str(today_date.day)
    if len(day) == 1:
        day = '0' + day

    return year+month+day


def copy_rename_template_file(template_file):

    # Create a date id from today's date
    date_id = create_date_id()

    # Assign original template file name and new template file name to variables
    new_file_name = 'Output/' + date_id + template_file
    old_file_name = 'Output/' + template_file

    # Check if a file exists in the target directory with the new file name, if yes delete it
    if os.path.isfile(new_file_name):
        os.remove(new_file_name)

    # Copy and rename original template file
    shutil.copy(template_file, 'Output')
    os.rename(old_file_name, new_file_name)

    return new_file_name


if __name__ == '__main__':
    template_file_input = 'xlwings_table_example.xlsx'
    csv_file_path_input = 'IMDB-Movie-Data.csv'

    data_from_csv = open_csv_file(csv_file_path_input)
    write_list_to_excel(template_file_input, data_from_csv)

