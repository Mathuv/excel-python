import pandas as pd
import numpy as np
import glob

all_data = pd.DataFrame()
for f in glob.glob("in\sales*.xlsx"):
    df = pd.read_excel(f)
    all_data = all_data.append(df, ignore_index=True)
print(all_data.describe())

pass