import re
str = "1ABC"
result = r'^1[A-Z]{2,3}$'
match = re.search(result, str)
if match:
    print("匹配成功:", match.group())
else:
    print("匹配失败")