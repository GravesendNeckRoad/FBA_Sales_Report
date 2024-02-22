import openpyxl as xl
import openpyxl.worksheet.worksheet as Worksheet


def data_bars(worksheet: Worksheet, column_number: int, color_hex: str) -> None:
    """Formats an excel column with data bars, based on relative min/max of said columns values.

    :param worksheet: (var: openpyxl.worksheet.worksheet ) The name of your openpyxl worksheet
    :param column_number: (int) the column number you wish to convert
    :param color_hex : (str) the hex code of the color you wish to assign to your bars (ex: '#FFFFFF')

    *IMPORTANT* Assumes that there is a 'totals' row, therefore starts at index 3. If your column doesn't have a subtotal
    or grand total, the data bars will not cover the full range. The report will still work, but the bars will be
    off by a bit. For example, the highest value will have the same sized bar as the second highest value.
    """

    column_letter = xl.utils.get_column_letter(column_number)
    col_for_databars = f"{column_letter}3:" \
                       f"{column_letter}{worksheet.max_row}"

    second_largest_value = \
        sorted(
            [cell.value for cell in worksheet[column_letter] if cell.coordinate != f"{column_letter}1"], reverse=True
        )[1]

    smallest_value = \
        sorted(
            [cell.value for cell in worksheet[column_letter] if cell.coordinate != f"{column_letter}1"]
        )[1]

    rule_on_hand = xl.formatting.rule.DataBarRule(
        start_type='num'
        , start_value=smallest_value
        , end_type='num'
        , end_value=second_largest_value
        , color=color_hex)

    # print(f'{second_largest_value} and {smallest_value}')

    worksheet.conditional_formatting.add(col_for_databars, rule_on_hand)

# todo
# errors to add later;
# ValueError
# "Please enter a NUMBER for your column that you wish to format. Do not use letters.
#
# NameError
# "There is no worksheet with this name. Please check to see what you are pulling from.
#
# ValueError("Colors must be aRGB hex values")
# "Colors must be aRGB hex values". "Blue" etc. are not accepted.
#
# AttributeError: 'int' object has no attribute 'max_row'
#
# TypeError: <class 'openpyxl.formatting.rule.DataBar'>.color should be <class 'openpyxl.styles.colors.Color'>
# but value is <class 'list'>
# "Please check your inputs. Lists/Dicts/etc. are not accepted. Use







