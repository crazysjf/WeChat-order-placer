import pandas as pd

import math
import os

dir = os.path.abspath(__file__)

print(dir)
dir = os.path.dirname(dir)
for i in range(0, 3):
    print(dir, dir.count('\\'))
    dir = os.path.dirname(dir)
