import os
import pandas as pd
import numpy as np

data_root = r"S:\lmassignment\LM_A2_2025_data"
result_dir = os.path.join(data_root, 'part1')
os.makedirs(result_dir, exist_ok=True)

if not os.path.exists(data_root):
    print(f"Error: Directory '{data_root}' not found.")
    exit()

subdirs = [d for d in os.listdir(data_root) if os.path.isdir(os.path.join(data_root, d)) and not d.startswith('part')]

for subdir in subdirs:
    subdir_path = os.path.join(data_root, subdir)
    xlsx_files = [f for f in os.listdir(subdir_path) if f.endswith(".xlsx")]
    
    file_types = {
        'food-C': 'food-C',
        'food-F': 'food-F',
        'money-C': 'money-C',
        'money-F': 'money-F'
    }
    
    processed_data = {}
    
    for xlsx in xlsx_files:
        full_path = os.path.join(subdir_path, xlsx)
        df = pd.read_excel(full_path)
        
        file_type = next((file_types[k] for k in file_types if k in xlsx), None)
        if file_type:
            df['SubjectiveValue'] = np.where(df['Var5'] == 0, df['Var1'], (df['Var1'] + df['Var2']) / 2)
            df['DiscountRate'] = np.where(
                (df['Var4'] > 0) & (df['SubjectiveValue'] > 0),
                (df['Var2'] - df['SubjectiveValue']) / (df['SubjectiveValue'] * df['Var4']),
                np.nan
            )
            processed_data[file_type] = df[['Var1', 'Var2', 'Var4', 'Var5', 'SubjectiveValue', 'DiscountRate']]
    
    for type_key, dataframe in processed_data.items():
        save_path = os.path.join(result_dir, f"{subdir}_{type_key}_discount_rates.csv")
        dataframe.to_csv(save_path, index=False)
        print(f"Saved {type_key} results to {save_path}")