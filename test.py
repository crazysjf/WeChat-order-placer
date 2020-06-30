import pandas as pd

import math
import os
import pandas

import constants
s = "11045-靓怡雅"
print(s in constants.PROVIDERS_IGNORING_DEFECTIVE_GOODS)


def _process(df):
    return df +1
ret = 1

ignored_defe_df = 2

(ret, ignored_defe_df) = map(_process, (ret, ignored_defe_df))

print(ret, ignored_defe_df)