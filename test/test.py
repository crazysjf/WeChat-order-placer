from datetime import date, datetime
t = date.today()   # 仅获取日期
print (t)

s = t.strftime("%m.%d")
print(s)
