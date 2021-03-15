import pandas as pd 
import numpy as np
import csv

df = pd.read_csv (r'2020-sd6.csv')
df.to_json (r'2020-sd6.json')

row_index = 0

df = pd.read_json (r'2020-sd6.json')
df.to_csv (r'2020-sd6-new.csv', index = None)

row_index = 0