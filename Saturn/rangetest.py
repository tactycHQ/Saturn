from openpyxl.utils import range_boundaries, get_column_letter
import itertools


# def range_boundaries(address):
#     # if this is normal reference then just use the openpyxl converter
#     boundaries = openpyxl_range_boundaries(address)
#     if None not in boundaries or ':' in address:
#         return boundaries


rn = 'A3:B5'
def computeRangeCells(rn):
    x = range_boundaries(rn)
    print(x)

    cols = list(range(min(x[0],x[2]), max(x[0],x[2])))
    cols.append(max(x[0],x[2]))
    rows = list(range(min(x[1],x[3]),max(x[1],x[3])))
    rows.append(max(x[1],x[3]))
    cols = list(map(get_column_letter, cols))

    final =[]
    for col in cols:
        for row in rows:
            final.append('{}{}'.format(col,row))
    print(final)
    return final

computeRangeCells(rn)


# for y2 in x2:
#     print(y2[0])
