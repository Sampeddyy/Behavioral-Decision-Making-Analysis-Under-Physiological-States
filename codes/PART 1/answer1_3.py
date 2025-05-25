import os
import pandas as pd
import numpy as np
from scipy.stats.mstats import gmean

root_dir = r"S:\lmassignment\LM_A2_2025_data"
target_dir = os.path.join(root_dir, 'part3')
os.makedirs(target_dir, exist_ok=True)

if not os.path.exists(root_dir):
    print(f"Error: Directory '{root_dir}' not found.")
    exit()

folder_list = [f for f in os.listdir(root_dir) if os.path.isdir(os.path.join(root_dir, f)) and not f.startswith('part')]

data_groups = ['food-C', 'food-F', 'money-C', 'money-F']

for folder in folder_list:
    folder_path = os.path.join(root_dir, folder)
    
    for group in data_groups:
        source_path = os.path.join(folder_path, 'part 2', f"{group}_switch_rates.xlsx")
        
        if not os.path.exists(source_path):
            continue
        
        data_frame = pd.read_excel(source_path)
        mean_values = data_frame['Geometric_Mean'].dropna().tolist()
        
        final_rate = gmean(mean_values) if mean_values else np.nan
        result_frame = pd.DataFrame({'Cumulative_Rate': [final_rate]})
        
        save_path = os.path.join(target_dir, f"{folder}_{group}_cumulative_rate.xlsx")
        result_frame.to_excel(save_path, index=False)
        print(f"Saved to {save_path}")