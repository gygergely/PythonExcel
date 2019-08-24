import os
import win32com.client as win32


def run_excel_macro (file_path, separator_char):
    """
    Execute an Excel macro
    :param file_path: path to the Excel file holding the macro
    :param separator_char: the character used by the operating system to separate pathname components
    :return: None
    """
    xl = win32.Dispatch('Excel.Application')
    xl.Application.visible = False

    try:
        wb = xl.Workbooks.Open(os.path.abspath(file_path))
        xl.Application.run(file_path.split(sep=separator_char)[-1] + "!main.simpleMain")
        wb.Save()
        wb.Close()

    except Exception as ex:
        template = "An exception of type {0} occurred. Arguments:\n{1!r}"
        message = template.format(type(ex).__name__, ex.args)
        print(message)

    xl.Application.Quit()
    del xl

separator_char = os.sep
run_excel_macro(input('Please enter Excel macro file path: '), separator_char)