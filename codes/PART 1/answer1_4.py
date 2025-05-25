import os
import pandas as pd
import numpy as np

base_dir = r"S:\lmassignment\LM_A2_2025_data"
output_dir = os.path.join(base_dir, 'part 4')
os.makedirs(output_dir, exist_ok=True)

if not os.path.exists(base_dir):
    print(f"Error: Directory '{base_dir}' not found.")
    exit()

subfolders = [s for s in os.listdir(base_dir) if os.path.isdir(os.path.join(base_dir, s))]

types = ['food-C', 'food-F', 'money-C', 'money-F']
rate_collection = {t: [] for t in types}

for sub in subfolders:
    sub_path = os.path.join(base_dir, sub)
    sub_output = os.path.join(sub_path, 'part 3')
    
    if not os.path.exists(sub_output):
        for t in types:
            rate_collection[t].append(np.nan)
        continue
    
    for t in types:
        file_path = os.path.join(sub_output, f"{t}_cumulative_rate.xlsx")
        if os.path.exists(file_path):
            data = pd.read_excel(file_path)
            rate_collection[t].append(data['Cumulative_Rate'].iloc[0])
        else:
            rate_collection[t].append(np.nan)

avg_rates = {t: np.nanmean(rate_collection[t]) for t in types}
result_data = pd.DataFrame({
    'Category': types,
    'Mean_Cumulative_Rate': [avg_rates[t] for t in types]
})

save_path = os.path.join(output_dir, 'mean_cumulative_rates.xlsx')
result_data.to_excel(save_path, index=False)
print(f"Saved to {save_path}")