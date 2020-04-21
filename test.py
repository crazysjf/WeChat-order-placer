import pandas as pd

import math


import datetime
detester = "2020/4/15 22:10:26"
date = datetime.datetime.strptime(detester,"%Y/%m/%d %H:%M:%S")
print(date)

now = datetime.datetime.now()
print(now)

print((now-date).total_seconds()/60/60/24)