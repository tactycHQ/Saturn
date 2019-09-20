from fastnumbers import fast_real

# def range_boundaries(address):
#     # if this is normal reference then just use the openpyxl converter
#     boundaries = openpyxl_range_boundaries(address)
#     if None not in boundaries or ':' in address:
#         return boundaries


x = '1'

y = fast_real(x)
print(type(y))


# for y2 in x2:
#     print(y2[0])
