import pandas as pd
import numpy as np
import glob

all_data = pd.DataFrame()
for f in glob.glob("in\sales*.xlsx"):
    df = pd.read_excel(f)
    all_data = all_data.append(df, ignore_index=True)

# debug
# print(all_data.describe())
# print(all_data.head())

#Conver the date column to a datetime object
all_data['date'] = pd.to_datetime(all_data['date'])

status = pd.read_excel("in/customer-status.xlsx")
# status
all_data_st = pd.merge(all_data, status, how='left')

#all_data_st.head()



pass