import pandas as pd
import shutil
import os

SRC_FOLDER = '.\\SRC_FILES'
SKIPPED_FOLDER = '.\\SKIPPED_FILES'
PROCESSED_FOLDER = '.\\PROCESSED_FILES'


def get_data_from_multiple_xl_files(src_folder_path):
    """
    Loading data from multiple Excel files to one pandas Dataframe.
    Assumptions:
    file extension is xlsx or xlsb (not tested with xls and xlsm)
    there is only one tab in the files
    file structures are the same (each file has a header in the first row, the nr of the columns are identical)
    :param src_folder_path: the source folder
    :return: dataframe
    """
    # sort the files alphabetically
    file_list = os.listdir(src_folder_path)
    file_list = sorted(file_list)

    file_counter = 0

    df_header = pd.DataFrame
    df_data = pd.DataFrame

    # iterate through all the files in the src folder
    for file in file_list:
        file = os.path.join(SRC_FOLDER, file)
        file_counter += 1
        file_extension = os.path.splitext(file)[1]

        # if file extension is xlsb or xlsx start reading data, else move to skipped folder
        if file_extension in ['.xlsb', '.xlsx']:
            if file_extension == '.xlsb':
                engine_str = 'pyxlsb'
            else:
                engine_str = None

            # get the "default" header from the first file and compare all the following files header to it
            # if there is no header (empty df) then move to the next one
            # if the 1st file header is incorrect all the other files move to skipped folder
            if file_counter == 1:
                df_header = pd.read_excel(file, engine=engine_str).columns
                if df_header.empty:
                    file_counter = 0
                    move_file_to_dir(file, SKIPPED_FOLDER)
                else:
                    df_data = pd.read_excel(file, engine=engine_str)
                    move_file_to_dir(file, PROCESSED_FOLDER)
            else:
                # if header is identical to the "default" header add data to data frame
                df_temp_header = pd.read_excel(file, engine=engine_str).columns
                if df_header.equals(df_temp_header):
                    df_data_to_add = pd.read_excel(file, engine=engine_str)
                    df_data = pd.concat([df_data, df_data_to_add], ignore_index=True)
                    move_file_to_dir(file, PROCESSED_FOLDER)
                else:
                    move_file_to_dir(file, SKIPPED_FOLDER)
        else:
            move_file_to_dir(file, SKIPPED_FOLDER)

    return df_data


def move_file_to_dir(file_name, destination_dir):
    """
    Move files from one folder to another folder, if the file already exists in the target folder delete it from
    target folder.
    :param file_name: absolute filepath of the file to be moved
    :param destination_dir: path of the target/destination folder
    :return: None
    """
    fn_name = file_name.split('\\')
    path_to_check = destination_dir + str(fn_name[-1])
    is_file_exists_in_destination = os.path.exists(path_to_check)
    if is_file_exists_in_destination:
        os.remove(path_to_check)
    shutil.move(file_name, destination_dir)


if __name__ == '__main__':
    data = get_data_from_multiple_xl_files(SRC_FOLDER)
    print(data.head(10))
    print(data.info())
    print(data.shape)
