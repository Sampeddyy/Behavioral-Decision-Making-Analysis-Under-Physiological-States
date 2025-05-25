import os
import pandas as pd
import numpy as np
from scipy.stats.mstats import gmean

base_path = r"S:\lmassignment\LM_A2_2025_data"
save_dir = os.path.join(base_path, 'part2')
os.makedirs(save_dir, exist_ok=True)

if not os.path.exists(base_path):
    print(f"Error: Directory '{base_path}' not found.")
    exit()

dirs = [d for d in os.listdir(base_path) if os.path.isdir(os.path.join(base_path, d)) and not d.startswith('part')]

data_types = ['food-C', 'food-F', 'money-C', 'money-F']

for current_dir in dirs:
    dir_path = os.path.join(base_path, current_dir)
    
    for data_type in data_types:
        source_file = os.path.join(dir_path, f"{data_type}_discount_rates.csv")
        
        if not os.path.exists(source_file):
            continue
        
        df_input = pd.read_csv(source_file)
        df_ordered = df_input.sort_values(by='k', ascending=False).reset_index(drop=True)
        
        switch_indices = df_ordered.index[df_ordered['Var5'].diff().abs() == 1].tolist()
        switch_combinations = [(switch_indices[j] - 1, switch_indices[j]) for j in range(len(switch_indices)) if switch_indices[j] > 0]
        
        switch_results = []
        for prev, curr in switch_combinations:
            rate_prev = abs(df_ordered.loc[prev, 'k'])
            rate_curr = abs(df_ordered.loc[curr, 'k'])
            mean_value = gmean([rate_prev, rate_curr]) if not np.isnan(rate_prev) and not np.isnan(rate_curr) else np.nan
            switch_results.append({
                'Trial_Prev': prev + 1,
                'Trial_Curr': curr + 1,
                'Geometric_Mean': mean_value
            })
        
        result_df = pd.DataFrame(switch_results)
        output_path = os.path.join(save_dir, f"{current_dir}_{data_type}_switch_rates.xlsx")
        result_df.to_excel(output_path, index=False)
        print(f"Saved to {output_path}")