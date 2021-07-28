import sys
import pandas as pd

data = pd.Series(['a','b','c'],index=[1,2,3])

print(data.max())

dataFrame = pd.DataFrame()