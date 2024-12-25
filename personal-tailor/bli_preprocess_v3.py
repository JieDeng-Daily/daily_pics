import pandas as pd
import glob
import sys, os

input_folder = sys.argv[1]
output_file = sys.argv[2]
file_paths = glob.glob(f"{input_folder}/*.txt") 

def merge_time(df):
    time_columns = [col for col in df.columns if 'Time' in col]
   # print(time_columns)
    if all(df[time_columns[0]].equals(df[col]) for col in time_columns[1:]):
        #print(all(df[time_columns[0]].equals(df[col]) for col in time_columns[1:]))
        df_time_merged = df.drop(columns=time_columns[1:])
    return df_time_merged
    
def rearrange_data(i,df):
    data_columns = [col for col in df.columns if 'Data' in col]
    fit_columns = [col for col in df.columns if 'Fit' in col]
    time_column = [col for col in df.columns if 'Time' in col]
    columns_all = time_column+data_columns+fit_columns
    columns_re = [col for col in columns_all if col in df.columns]
    #print(columns_re)
    df_sorted = df[columns_re]
    #print(df_sorted)
    return df_sorted 

def normalize_time(i,df):
    time_column = [col for col in df.columns if 'Time' in col]
    df[time_column] = df[time_column]-df[time_column].min()
    return df

data_by_time = {}

for i, file_path in enumerate(file_paths, start=1):
    #prefix = f"Cycle{i}"
    data = pd.read_csv(file_path, sep='\t', skiprows=range(0,7))
    file_name = os.path.basename(file_path) 
    if "Results.txt" in file_name:
        sensor_name = file_name[:2]
        #print(sensor_name)
    data = data.loc[:, ~data.columns.str.contains('^Unnamed')]
    
    for col in data.columns:
        data_type = col.rstrip('1234567890') 
        #print(data_type)
        time_index = ''.join(filter(str.isdigit, col))
        #print(time_index)
        group_name = f"{data_type}{time_index}_{sensor_name}"
        #print(group_name)

        if time_index not in data_by_time:
            data_by_time[time_index] = pd.DataFrame()
        
        data_by_time[time_index][group_name] = data[col]
        
merged_data = pd.concat(data_by_time.values(), axis=1)
#merged_data.to_csv('1.csv')

empty_columns=merged_data.columns[merged_data.isnull().all()]
dfs = []
last_index = 0

dfs = []
n = 15  # 每个子数据框的列数

for i in range(0, len(merged_data.columns), n):
    split_df = merged_data.iloc[:, i:i + n]
    dfs.append(split_df)

with pd.ExcelWriter(output_file) as writer:
    
    for i, df_part in enumerate(dfs, 1):
        #print(df_part)

        df_time_merged = merge_time(df_part)
        #print(df_time_merged)
        df_sorted = rearrange_data(i, df_time_merged)
        df_time_normalized = normalize_time(i, df_sorted)
        
        sheet_name = f'Sheet_{i}' 
        df_time_normalized.to_excel(writer, sheet_name=sheet_name, index=False)
