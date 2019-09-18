

def asum(a,b):
    return a+b

ret = lambda x,y: asum(x,y)
lam = 'lambda x,y: asum(x,y)'

x = eval(lam)(2,2)
print(x)

y = '2+3+4'
print(eval(y))