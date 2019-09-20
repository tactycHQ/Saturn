from fastnumbers import fast_real

# def range_boundaries(address):
#     # if this is normal reference then just use the openpyxl converter
#     boundaries = openpyxl_range_boundaries(address)
#     if None not in boundaries or ':' in address:
#         return boundaries

x= {
    'a':[1,2,3],
    'b':[4,5,6]
}

y = 1

if y not in x.get('a'):
    x['a'].append(y)

print(x)


# for y2 in x2:
#     print(y2[0])
