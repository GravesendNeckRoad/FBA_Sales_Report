import openpyxl as xl
from openpyxl.worksheet.worksheet import Worksheet


def align_and_center(worksheet: Worksheet) -> None:
    """ adjust column width based on text, then center said text, of your excel data array

    :param worksheet: (var) your openpyxl worksheet
    """
    # auto-fit column width
    for column in worksheet.columns:
        column_letter = xl.utils.get_column_letter(column[0].column)
        max_length = 0

        for cell in column:
            if len(str(cell.value)) > max_length:
                max_length = len(str(cell.value))
            worksheet.column_dimensions[column_letter].width = max_length + 5  # add a little extra padding

            # center every cell, while we're at it
            cell.alignment = xl.styles.Alignment(
                horizontal='center'
                , vertical='center'
            )
