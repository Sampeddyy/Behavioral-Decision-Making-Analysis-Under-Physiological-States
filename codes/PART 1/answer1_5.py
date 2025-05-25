import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

root_path = r"S:\lmassignment\LM_A2_2025_data"
save_path = os.path.join(root_path, 'part 5')
os.makedirs(save_path, exist_ok=True)

avg_data = pd.read_excel(os.path.join(root_path, 'part 4', 'mean_cumulative_rates.xlsx'))
avg_dict = {row['Category']: row['Mean_Cumulative_Rate'] for _, row in avg_data.iterrows()}

folder_list = [f for f in os.listdir(root_path) if os.path.isdir(os.path.join(root_path, f))]
data_types = ['food-C', 'food-F', 'money-C', 'money-F']
indiv_rates = {t: [] for t in data_types}

for folder in folder_list:
    for t in data_types:
        file_path = os.path.join(root_path, folder, 'part 3', f"{t}_cumulative_rate.xlsx")
        indiv_rates[t].append(pd.read_excel(file_path)['Cumulative_Rate'].iloc[0] if os.path.exists(file_path) else np.nan)

food_rel_diff = (avg_dict['food-F'] - avg_dict['food-C']) / avg_dict['food-C']
money_rel_diff = (avg_dict['money-F'] - avg_dict['money-C']) / avg_dict['money-C']

food_indiv_diffs = [(indiv_rates['food-F'][i] - indiv_rates['food-C'][i]) / indiv_rates['food-C'][i] if not np.isnan(indiv_rates['food-C'][i]) and indiv_rates['food-C'][i] != 0 else np.nan for i in range(len(folder_list))]
money_indiv_diffs = [(indiv_rates['money-F'][i] - indiv_rates['money-C'][i]) / indiv_rates['money-C'][i] if not np.isnan(indiv_rates['money-C'][i]) and indiv_rates['money-C'][i] != 0 else np.nan for i in range(len(folder_list))]
food_sem = np.nanstd(food_indiv_diffs) / np.sqrt(np.sum(~np.isnan(food_indiv_diffs)))
money_sem = np.nanstd(money_indiv_diffs) / np.sqrt(np.sum(~np.isnan(money_indiv_diffs)))

plt.figure(figsize=(6, 4))
plot_bars = plt.bar(['Food', 'Money'], [food_rel_diff, money_rel_diff], yerr=[food_sem, money_sem], capsize=5, color=['#FF9999', '#66B2FF'], edgecolor='black')
plt.axhline(0, color='gray', linestyle='--')
for bar in plot_bars:
    height = bar.get_height()
    plt.text(bar.get_x() + bar.get_width()/2, height + (0.01 if height >= 0 else -0.02), f'{height:.3f}', ha='center', va='bottom' if height >= 0 else 'top')
plt.tight_layout()
plt.savefig(os.path.join(save_path, 'relative_difference_plot.png'))
plt.close()

result_df = pd.DataFrame({'Category': ['Food', 'Money'], 'Relative_Difference': [food_rel_diff, money_rel_diff], 'SEM': [food_sem, money_sem]})
result_df.to_excel(os.path.join(save_path, 'relative_differences.xlsx'), index=False)
print(f"Saved to {os.path.join(save_path, 'relative_differences.xlsx')}")