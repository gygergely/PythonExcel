import xlwings as xw


def run_excel_macro(file_path):
    """
    Execute an Excel macro
    :param file_path: path to the Excel file holding the macro
    :return: None
    """

    try:
        xl_app = xw.App(visible=False, add_book=False)
        wb = xl_app.books.open(file_path)

        run_macro = wb.app.macro('main.SimpleMain')
        run_macro()

        wb.save()
        wb.close()

        xl_app.quit()

    except Exception as ex:
        template = "An exception of type {0} occurred. Arguments:\n{1!r}"
        message = template.format(type(ex).__name__, ex.args)
        print(message)


run_excel_macro(input('Please enter Excel macro file path: '))
