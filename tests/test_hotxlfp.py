import hotxlfp
p = hotxlfp.Parser()

result = p.parse('SUM({1,5}*{3,7})')
print(result)