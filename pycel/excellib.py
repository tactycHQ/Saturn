"""
Python equivalents of various excel functions
"""
import math
from bisect import bisect_right
from decimal import Decimal, ROUND_HALF_UP, ROUND_UP

import numpy as np

from pycel.excelutil import (
    build_wildcard_re,
    DIV0,
    ERROR_CODES,
    ExcelCmp,
    find_corresponding_index,
    flatten,
    handle_ifs,
    is_number,
    list_like,
    MAX_COL,
    MAX_ROW,
    NA_ERROR,
    NUM_ERROR,
    REF_ERROR,
    VALUE_ERROR,
)
from pycel.lib.function_helpers import (
    excel_helper,
    excel_math_func,
)


def _numerics(*args, keep_bools=False):
    # ignore non numeric cells
    args = tuple(flatten(args))
    error = next((x for x in args if x in ERROR_CODES), None)
    if error is not None:
        # return the first error in the list
        return error
    else:
        if not keep_bools:
            args = (a for a in args if not isinstance(a, bool))
        return tuple(x for x in args if isinstance(x, (int, float)))


def average(*args):
    data = _numerics(*args)

    # A returned string is an error code
    if isinstance(data, str):
        return data
    elif len(data) == 0:
        return DIV0
    else:
        return sum(data) / len(data)


def averageifs(average_range, *args):
    # Excel reference: https://support.office.com/en-us/article/
    #   AVERAGEIFS-function-48910C45-1FC0-4389-A028-F7C5C3001690

    coords = handle_ifs(args, average_range)
    data = _numerics((average_range[r][c] for r, c in coords), keep_bools=True)
    if len(data) == 0:
        return DIV0
    return sum(data) / len(data)


@excel_math_func
def ceiling(number, significance):
    # Excel reference: https://support.office.com/en-us/article/
    #   CEILING-function-0A5CD7C8-0720-4F0A-BD2C-C943E510899F
    if significance < 0 < number:
        return NUM_ERROR

    if number == 0:
        return 0

    if significance == 0:
        return DIV0

    if number < 0 < significance:
        return significance * int(number / significance)
    else:
        return significance * math.ceil(number / significance)


@excel_helper()
def column(ref):
    # Excel reference: https://support.office.com/en-us/article/
    #   COLUMN-function-44E8C754-711C-4DF3-9DA4-47A55042554B

    if ref.is_range:
        if ref.end.col_idx == 0:
            return range(1, MAX_COL + 1)
        else:
            return (tuple(range(ref.start.col_idx, ref.end.col_idx + 1)), )
    else:
        return ref.col_idx


def count(*args):
    # Excel reference: https://support.office.com/en-us/article/
    #   COUNT-function-a59cd7fc-b623-4d93-87a4-d23bf411294c

    total = 0

    for arg in args:
        if isinstance(arg, list):
            # count inside a list
            total += len(
                [x for x in arg if is_number(x) and not isinstance(x, bool)])
        else:
            total += int(is_number(arg))

    return total


def countif(range, criteria):
    # Excel reference: https://support.office.com/en-us/article/
    #   COUNTIF-function-e0de10c6-f885-4e71-abb4-1f464816df34

    # WARNING:
    # - wildcards not supported  ::TODO:: test if this is no longer true
    # - support of strings with >, <, <=, =>, <> not provided

    valid = find_corresponding_index(range, criteria)

    return len(valid)


def conditional_format_ids(*args):
    """helper function for getting conditional format ids"""
    # Excel reference: https://support.office.com/en-us/article/
    #   E09711A3-48DF-4BCB-B82C-9D8B8B22463D

    results = []
    for condition, dxf_id, stop_if_true in args:
        if condition:
            results.append(dxf_id)
            if stop_if_true:
                break
    return results


def countifs(*args):
    # Excel reference: https://support.office.com/en-us/article/
    #   COUNTIFS-function-dda3dc6e-f74e-4aee-88bc-aa8c2a866842

    return len(handle_ifs(args))


@excel_math_func
def floor(number, significance):
    # Excel reference: https://support.office.com/en-us/article/
    #   FLOOR-function-14BB497C-24F2-4E04-B327-B0B4DE5A8886
    if significance < 0 < number or number < 0 < significance:
        return NUM_ERROR

    if number == 0:
        return 0

    if significance == 0:
        return DIV0

    return significance * int(number / significance)


@excel_helper(cse_params=0, bool_params=3, number_params=2)
def hlookup(lookup_value, table_array, row_index_num, range_lookup=True):
    """ Horizontal Lookup

    :param lookup_value: value to match (value or cell reference)
    :param table_array: range of cells being searched.
    :param row_index_num: column number to return
    :param range_lookup: True, assumes sorted, finds nearest. False: find exact
    :return: #N/A if not found else value
    """
    # Excel reference: https://support.office.com/en-us/article/
    #   hlookup-function-a3034eec-b719-4ba3-bb65-e1ad662ed95f

    if not list_like(table_array):
        return NA_ERROR

    if row_index_num <= 0:
        return VALUE_ERROR

    if row_index_num > len(table_array[0]):
        return REF_ERROR

    result_idx = _match(
        lookup_value, table_array[0], match_type=bool(range_lookup))

    if isinstance(result_idx, int):
        return table_array[row_index_num - 1][result_idx - 1]
    else:
        # error string
        return result_idx


@excel_helper(number_params=(1, 2))
def index(array, row_num, col_num=None):
    # Excel reference: https://support.office.com/en-us/article/
    #   index-function-a5dcf0dd-996d-40a4-a822-b56b061328bd

    if not list_like(array) or not list_like(array[0]):
        return VALUE_ERROR

    try:
        # rectangular array
        if row_num and col_num:
            return array[row_num - 1][col_num - 1]

        elif row_num:
            if len(array[0]) == 1:
                return array[row_num - 1][0]
            elif len(array) == 1:
                return array[0][row_num - 1]
            elif isinstance(array, np.ndarray):
                return array[row_num - 1, :]
            else:
                return (tuple(array[row_num - 1]),)

        elif col_num:
            if len(array) == 1:
                return array[0][col_num - 1]
            elif len(array[0]) == 1:
                return array[col_num - 1][0]
            elif isinstance(array, np.ndarray):
                result = array[:, col_num - 1]
                result.shape = result.shape + (1,)
                return result
            else:
                return tuple((r[col_num - 1], ) for r in array)

    except IndexError:
        pass

    return NA_ERROR


@excel_helper(cse_params=0, err_str_params=None)
def iserr(value):
    # Excel reference: https://support.office.com/en-us/article/
    #   is-functions-0f2d7971-6019-40a0-a171-f2d869135665
    # Value refers to any error value except #N/A.
    return isinstance(value, str) and value in ERROR_CODES and value != NA_ERROR


@excel_helper(cse_params=0, err_str_params=None)
def iserror(value):
    # Excel reference: https://support.office.com/en-us/article/
    #   is-functions-0f2d7971-6019-40a0-a171-f2d869135665
    # Value refers to any error value:
    #   (#N/A, #VALUE!, #REF!, #DIV/0!, #NUM!, #NAME?, or #NULL!).
    return isinstance(value, str) and value in ERROR_CODES or (
        isinstance(value, tuple))


@excel_helper(cse_params=0, err_str_params=None)
def istext(arg):
    # Excel reference: https://support.office.com/en-us/article/
    #   is-functions-0f2d7971-6019-40a0-a171-f2d869135665
    return isinstance(arg, str) and arg not in ERROR_CODES


@excel_helper(cse_params=0, err_str_params=None)
def isna(value):
    # Excel reference: https://support.office.com/en-us/article/
    #   is-functions-0f2d7971-6019-40a0-a171-f2d869135665
    return value == NA_ERROR or isinstance(value, tuple)


@excel_helper(cse_params=0, err_str_params=None)
def isnumber(value):
    # Excel reference: https://support.office.com/en-us/article/
    #   is-functions-0f2d7971-6019-40a0-a171-f2d869135665
    return isinstance(value, (int, float))


def linest(Y, X, const=True, degree=1):  # pragma: no cover  ::TODO::
    if isinstance(const, str):
        const = (const.lower() == "true")

    def assert_vector(data):
        vector = np.array(data)
        assert 1 in vector.shape
        return vector.ravel()

    X = assert_vector(X)
    Y = assert_vector(Y)

    # build the vandermonde matrix
    A = np.vander(X, degree + 1)

    if not const:
        # force the intercept to zero
        A[:, -1] = np.zeros((1, len(X)))

    # perform the fit
    coefs, residuals, rank, sing_vals = np.linalg.lstsq(A, Y, rcond=None)

    return coefs


@excel_math_func
def ln(arg):
    # Excel reference: https://support.office.com/en-us/article/
    #   LN-function-81FE1ED7-DAC9-4ACD-BA1D-07A142C6118F
    return math.log(arg)


@excel_math_func
def log(number, base=10):
    # Excel reference: https://support.office.com/en-us/article/
    #   LOG-function-4E82F196-1CA9-4747-8FB0-6C4A3ABB3280
    return math.log(number, base)


@excel_helper(cse_params=0)
def lookup(lookup_value, lookup_array, result_range=None):
    """
    There are two ways to use LOOKUP: Vector form and Array form

    Vector form: lookup_array is list like (ie: n x 1)

    Array form: lookup_array is rectangular (ie: n x m)

        First row or column is the lookup vector.
        Last row or column is the result vector
        The longer dimension is the search dimension

    :param lookup_value: value to match (value or cell reference)
    :param lookup_array: range of cells being searched.
    :param result_range: (optional vector form) values are returned from here
    :return: #N/A if not found else value
    """
    if not list_like(lookup_array):
        return NA_ERROR

    height = len(lookup_array)
    width = len(lookup_array[0])

    # match across the largest dimension
    if width <= height:
        match_idx = _match(lookup_value, tuple(i[0] for i in lookup_array))
        result = tuple(i[-1] for i in lookup_array)
    else:
        match_idx = _match(lookup_value, lookup_array[0])
        result = lookup_array[-1]

    if len(lookup_array) > 1 and len(lookup_array[0]) > 1:
        # rectangular array
        assert result_range is None

    elif result_range:
        if len(result_range) > len(result_range[0]):
            result = tuple(i[0] for i in result_range)
        else:
            result = result_range[0]

    if isinstance(match_idx, int):
        return result[match_idx - 1]

    else:
        # error string
        return match_idx


@excel_helper(cse_params=0, number_params=2)
def match(lookup_value, lookup_array, match_type=1):
    if len(lookup_array) == 1:
        lookup_array = lookup_array[0]
    else:
        lookup_array = tuple(row[0] for row in lookup_array)

    return _match(lookup_value, lookup_array, match_type)


def _match(lookup_value, lookup_array, match_type=1):
    # Excel reference: https://support.office.com/en-us/article/
    #   MATCH-function-E8DFFD45-C762-47D6-BF89-533F4A37673A

    """ The relative position of a specified item in a range of cells.

    Match_type Behavior

    1: return the largest value that is less than or equal to
    `lookup_value`. `lookup_array` must be in ascending order.

    0: return the first value that is exactly equal to lookup_value.
    `lookup_array` can be in any order.

    -1: return the smallest value that is greater than or equal to
    `lookup_value`. `lookup_array` must be in descending order.

    If `match_type` is 0 and lookup_value is a text string, you can use the
    wildcard characters — the question mark (?) and asterisk (*).

    :param lookup_value: value to match (value or cell reference)
    :param lookup_array: range of cells being searched.
    :param match_type: The number -1, 0, or 1.
    :return: #N/A if not found, or relative position in `lookup_array`
    """
    lookup_value = ExcelCmp(lookup_value)

    if match_type == 1:
        # Use a binary search to speed it up.  Excel seems to do this as it
        # would explain the results seen when doing out of order searches.
        lookup_value = ExcelCmp(lookup_value)

        result = bisect_right(lookup_array, lookup_value)
        while result and lookup_value.cmp_type != ExcelCmp(
                lookup_array[result - 1]).cmp_type:
            result -= 1
        if result == 0:
            result = NA_ERROR
        return result

    result = [NA_ERROR]

    if match_type == 0:
        def compare(idx, val):
            if val == lookup_value:
                result[0] = idx
                return True

        if lookup_value.cmp_type == 1:
            # string matches might be wildcards
            re_compare = build_wildcard_re(lookup_value.value)
            if re_compare is not None:
                def compare(idx, val):  # noqa: F811
                    if re_compare(val.value):
                        result[0] = idx
                        return True
    else:
        def compare(idx, val):
            if val < lookup_value:
                return True
            result[0] = idx
            return val == lookup_value

    for i, value in enumerate(lookup_array, 1):
        if value not in ERROR_CODES:
            value = ExcelCmp(value)
            if value.cmp_type == lookup_value.cmp_type and compare(i, value):
                break

    return result[0]


@excel_math_func
def mod(number, divisor):
    # Excel reference: https://support.office.com/en-us/article/
    #   MOD-function-9b6cd169-b6ee-406a-a97b-edf2a9dc24f3
    if divisor == 0:
        return DIV0

    return number % divisor


@excel_math_func
def npv(*args):
    # Excel reference: https://support.office.com/en-us/article/
    #   NPV-function-8672CB67-2576-4D07-B67B-AC28ACF2A568

    rate = args[0] + 1
    cashflow = args[1:]
    return sum([float(x) * rate ** -i for i, x in enumerate(cashflow, start=1)])


@excel_math_func
def power(number, power):
    # Excel reference: https://support.office.com/en-us/article/
    #   POWER-function-D3F2908B-56F4-4C3F-895A-07FB519C362A
    if number == power == 0:
        # Really excel?  What were you thinking?
        return NA_ERROR

    try:
        return number ** power
    except ZeroDivisionError:
        return DIV0


@excel_math_func
def roundup(number, num_digits):
    # Excel reference: https://support.office.com/en-us/article/
    #   ROUNDUP-function-F8BC9B23-E795-47DB-8703-DB171D0C42A7
    num_digits = int(num_digits)
    quant = Decimal('1E{}{}'.format('+-'[num_digits >= 0], abs(num_digits)))
    return float(Decimal(repr(number)).quantize(quant, rounding=ROUND_UP))


@excel_helper()
def row(ref):
    # Excel reference: https://support.office.com/en-us/article/
    #   ROW-function-3A63B74A-C4D0-4093-B49A-E76EB49A6D8D
    if ref.is_range:
        if ref.end.row == 0:
            return range(1, MAX_ROW + 1)
        else:
            return tuple((c, ) for c in range(ref.start.row, ref.end.row + 1))
    else:
        return ref.row


def sumif(rng, criteria, sum_range=None):
    # Excel reference: https://support.office.com/en-us/article/
    #   SUMIF-function-169b8c99-c05c-4483-a712-1697a653039b

    # WARNING:
    # - The following is not currently implemented:
    #  The sum_range argument does not have to be the same size and shape as
    #  the range argument. The actual cells that are added are determined by
    #  using the upper leftmost cell in the sum_range argument as the
    #  beginning cell, and then including cells that correspond in size and
    #  shape to the range argument.

    if sum_range is None:
        sum_range = rng
    return sumifs(sum_range, rng, criteria)


def sumifs(sum_range, *args):
    # Excel reference: https://support.office.com/en-us/article/
    #   SUMIFS-function-C9E748F5-7EA7-455D-9406-611CEBCE642B

    return sum(_numerics(
        (sum_range[r][c] for r, c in handle_ifs(args, sum_range)),
        keep_bools=True
    ))


def sumproduct(*args):
    # Excel reference: https://support.office.com/en-us/article/
    #   SUMPRODUCT-function-16753E75-9F68-4874-94AC-4D2145A2FD2E

    # find any errors
    error = next((i for i in flatten(args) if i in ERROR_CODES), None)
    if error:
        return error

    # verify array sizes match
    sizes = set()
    for arg in args:
        assert isinstance(arg, tuple), isinstance(arg[0], tuple)
        sizes.add((len(arg), len(arg[0])))
    if len(sizes) != 1:
        return VALUE_ERROR

    # put the values into numpy vectors
    values = np.array(tuple(tuple(
        x if isinstance(x, (float, int)) and not isinstance(x, bool) else 0
        for x in flatten(arg)) for arg in args))

    # return the sum product
    return np.sum(np.prod(values, axis=0))


@excel_math_func
def trunc(number, num_digits=0):
    # Excel reference: https://support.office.com/en-us/article/
    #   TRUNC-function-8B86A64C-3127-43DB-BA14-AA5CEB292721
    factor = 10 ** int(num_digits)
    return int(number * factor) / factor


@excel_helper(cse_params=0, bool_params=3, number_params=2)
def vlookup(lookup_value, table_array, col_index_num, range_lookup=True):
    """ Vertical Lookup

    :param lookup_value: value to match (value or cell reference)
    :param table_array: range of cells being searched.
    :param col_index_num: column number to return
    :param range_lookup: True, assumes sorted, finds nearest. False: find exact
    :return: #N/A if not found else value
    """
    # Excel reference: https://support.office.com/en-us/article/
    #   VLOOKUP-function-0BBC8083-26FE-4963-8AB8-93A18AD188A1

    if not list_like(table_array):
        return NA_ERROR

    if col_index_num <= 0:
        return '#VALUE!'

    if col_index_num > len(table_array[0]):
        return REF_ERROR

    result_idx = _match(
        lookup_value,
        [row[0] for row in table_array],
        match_type=bool(range_lookup)
    )

    if isinstance(result_idx, int):
        return table_array[result_idx - 1][col_index_num - 1]
    else:
        # error string
        return result_idx


@excel_math_func
def x_abs(value1):
    # Excel reference: https://support.office.com/en-us/article/
    #   ABS-function-3420200F-5628-4E8C-99DA-C99D7C87713C
    return abs(value1)


@excel_math_func
def xatan2(x_num, y_num):
    # Excel reference: https://support.office.com/en-us/article/
    #   ATAN2-function-C04592AB-B9E3-4908-B428-C96B3A565033

    # swap arguments
    return math.atan2(y_num, x_num)


@excel_math_func
def x_int(value1):
    # Excel reference: https://support.office.com/en-us/article/
    #   INT-function-A6C4AF9E-356D-4369-AB6A-CB1FD9D343EF
    return math.floor(value1)


def xmax(*args):
    data = _numerics(*args)

    # A returned string is an error code
    if isinstance(data, str):
        return data

    # however, if no non numeric cells, return zero (is what excel does)
    elif len(data) < 1:
        return 0
    else:
        return max(data)


def xmin(*args):
    data = _numerics(*args)

    # A returned string is an error code
    if isinstance(data, str):
        return data

    # however, if no non numeric cells, return zero (is what excel does)
    elif len(data) < 1:
        return 0
    else:
        return min(data)


@excel_math_func
def x_round(number, num_digits=0):
    # Excel reference: https://support.office.com/en-us/article/
    #   ROUND-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c

    num_digits = int(num_digits)
    if num_digits >= 0:  # round to the right side of the point
        return float(Decimal(repr(number)).quantize(
            Decimal(repr(pow(10, -num_digits))),
            rounding=ROUND_HALF_UP
        ))
        # see https://docs.python.org/2/library/functions.html#round
        # and https://gist.github.com/ejamesc/cedc886c5f36e2d075c5

    else:
        return round(number, num_digits)


def xsum(*args):
    data = _numerics(*args)
    if isinstance(data, str):
        return data

    # if no non numeric cells, return zero (is what excel does)
    return sum(data)
