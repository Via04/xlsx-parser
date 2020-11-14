from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.exceptions import InvalidFileException


# Static Service functions
def get_number(st: str) -> tuple:
    """
    Checks if given parameter is number, returns true if it is (number can be divided by comma or by point) returns True
    if number and number in python format as second argument if it is not number returns None
    :param st: str
    :return: bool
    """
    try:
        tmp = float(st)
        return True, tmp
    except (ValueError, TypeError):
        try:
            list_num = st.split(',')
            if len(list_num) == 2:
                neg = False
                if list_num[0].startswith('-'):
                    neg = True
                    list_num[0] = list_num[0][1:]
                if list_num[0].isdigit() and list_num[1].isdigit():
                    if neg:
                        return True, -1 * (float(list_num[0] + '.' + list_num[1]))
                    else:
                        return True, float(list_num[0] + '.' + list_num[1])
            else:
                return False, None
        except AttributeError:
            return False, None
        # print(str(st) + ' is not a number')
        return False, None


def trunc(num: float, precision: int) -> float:
    """
    Returns value truncated to precision signs after comma
    :param precision: int
    :param num: float
    :return: float
    """
    formation = '{:.' + str(precision) + 'f}'
    return float(formation.format(num))


def is_integer(num) -> bool:
    """
    Checks if given number is integer
    :param num: int, float
    :return: bool
    """
    return int(num) == float(num)


def precision_num(num: float) -> int:
    """
    Finds number of signs after comma in a float number, in case of error returns -1, in case of integer -2
    :param num: float
    :return: int
    """
    try:
        s = str(num)
        if '.' in s:
            return len(s) - s.find('.') - 1
        else:
            print('number is not float')
            return -2
    except(ValueError, TypeError):
        print('error parsing', -1)
        return -1


def is_formula(st: str) -> bool:
    """
    checks if given parameter is formula from exel
    :param st: str
    :return: bool
    """
    return True if st.startswith('=') else False


# Main Class
def get_end_row(col: tuple, start: int) -> int:
    """
    finds ending row in given column and sums it with start - 1 (to get last row)
    :param col: tuple
    :param start: int
    :return: int
    """
    counter = 0
    while col[counter] is not None and col[counter] != '' and counter < len(col) - 1:
        counter += 1
    return start + counter - 1


class XLSXParser:
    STARTING_ROW = 5  # the number of row after headers
    ending_row = 600
    filepath = 'default_name.xlsx'
    _wb = None  # variable to store .xlsx Workbook
    _ws = None  # variable to store worksheet
    red_fill = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')  # default error color (RED)

    def __init__(self, path: str):
        try:
            self.filepath = path  # writes path to save
            self._wb = load_workbook(path)  # loads .xlsx file
            self._wb.active = 0  # sets first sheet as active
            self._ws = self._wb.active  # a reference to a worksheet
            col = self._ws.iter_cols(min_col=3,
                                     max_col=3,
                                     min_row=self.STARTING_ROW,
                                     max_row=self.ending_row,
                                     values_only=True)
            self.ending_row = get_end_row(next(col), self.STARTING_ROW)
            print(self.ending_row)
        except (InvalidFileException, FileNotFoundError):
            print("Error! Bad path!")

    def fix_column(self, col: tuple, col_num: int):
        """
        Fixes and marks errors by checking for number cell of a particular column, workbook variable must be defined
        :param col: tuple
        :param col_num: int
        :return: None
        """
        if self._ws is not None:  # check is worksheet exists
            row_counter = self.STARTING_ROW  # initializing counter on the first row after header
            for cel in col:
                if cel is not None:  # checking cell conditions
                    is_num, num = get_number(cel)
                    if is_num:
                        # print(str(cel) + ' is a number')
                        if is_integer(num):
                            # print(str(cel) + ' is integer')
                            self._ws.cell(row=row_counter, column=col_num, value=num)  # if it is int - do nothing
                        else:  # but we still assign the value to fix possible error when number has wrong separator and
                            # print(str(cel) + ' is float')
                            if precision_num(num) > 5:  # get_number() function fixes this problem
                                self._ws.cell(row=row_counter, column=col_num, value=trunc(num, 5))  # if it is float
                            else:                                   # and contains many signs after comma trunc value
                                self._ws.cell(row=row_counter, column=col_num, value=num)  # else do not touch
                    else:
                        if not is_formula(cel):
                            # print(str(cel) + 'contains error')
                            self._ws.cell(row=row_counter, column=col_num).fill = self.red_fill  # and if it is not
                    row_counter += 1  # a number at all mark it with marking color.

    def find_errors(self):  # this function basically used just to parse through the table
        if self._ws is not None:
            col_counter = 4
            subsection_amount = 961
            num_subsections = ['Раз', 'Объем', 'Расценка', 'Годовая стоимость']
            text_subsections = ['Периодичность', 'Ед.изм.']
            for i in range(subsection_amount):
                # print(self._ws.cell(column=col_counter, row=self.STARTING_ROW - 1).value)
                if self._ws.cell(column=col_counter, row=self.STARTING_ROW - 1).value in num_subsections:
                    col = self._ws.iter_cols(min_col=col_counter,
                                             max_col=col_counter,
                                             min_row=self.STARTING_ROW,
                                             max_row=self.ending_row,
                                             values_only=True)
                    col_tup = next(col)
                    self.fix_column(col_tup, col_counter)
                    # print(col_tup)
                col_counter += 1

            self._wb.save(self.filepath)


s = input('input path to .xlsx file: ')
# print(s)
xl = XLSXParser(s)
xl.find_errors()
