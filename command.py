import re

string = 'ALLIED  CSE305 Parallel  Computing from 11:35 AM to 12:25 PM'
pattern = '[0-9]{1,2}:[0-9]{1,2}'

result = re.findall(pattern, string)

print(result,type(result[1]))
a=result[0]
b=result[1]
print(a,b)
