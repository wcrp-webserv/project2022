import numpy as np 
import pandas as pd 


data = np.array([['1', '2'], ['3', '4']]) 

dataFrame = pd.DataFrame(data, columns = ['col1', 'col2']) 
json = dataFrame.to_json() 
print(json) 

json_split = dataFrame.to_json(orient ='split') 
print("json_split = ", json_split, "\n") 

json_records = dataFrame.to_json(orient ='records') 
print("json_records = ", json_records, "\n") 

json_index = dataFrame.to_json(orient ='index') 
print("json_index = ", json_index, "\n") 

json_columns = dataFrame.to_json(orient ='columns') 
print("json_columns = ", json_columns, "\n") 

json_values = dataFrame.to_json(orient ='values') 
print("json_values = ", json_values, "\n") 

json_table = dataFrame.to_json(orient ='table') 
print("json_table = ", json_table, "\n") 

