import os
import pandas as pd
import numpy as np
from scipy.stats import gmean
import matplotlib.pyplot as plt

# Main folder
main_folder = r"S:\lmassignment\LM_A2_2025_data"
part2_2_folder = os.path.join(main_folder, 'part2_2')
os.makedirs(part2_2_folder, exist_ok=True)

# Step 1: Load and Filter Data
participants = [f for f in os.listdir(main_folder) if os.path.isdir(os.path.join(main_folder, f)) and not f.startswith('part')]
categories = {
    'Small_Food': ('food-F', lambda x: 1 <= x <= 5),
    'Large_Food': ('food-F', lambda x: 6 <= x <= 10),
    'Small_Money': ('money-F', lambda x: 1 <= x <= 10),
    'Large_Money': ('money-F', lambda x: 11 <= x <= 20)
}

data = {cat: {} for cat in categories}
for p in participants:
    for cat, (file_cat, filt) in categories.items():
        file = os.path.join(main_folder, p, f"{p}-{file_cat}.xlsx")
        if os.path.exists(file):
            df = pd.read_excel(file)
            filtered_df = df[df['Var1'].apply(filt)].copy()
            data[cat][p] = filtered_df
        else:
            data[cat][p] = pd.DataFrame()

# Step 2: Compute Discount Rate (k)
for cat in data:
    for p in data[cat]:
        df = data[cat][p]
        if not df.empty:
            df['V'] = np.where(df['Var5'] == 0, df['Var1'], (df['Var1'] + df['Var2']) / 2)
            df['k'] = np.where((df['Var4'] > 0) & (df['V'] > 0), (df['Var2'] - df['V']) / (df['V'] * df['Var4']), np.nan)

# Step 3: Compute Cumulative Rates
cumulative_rates = {cat: {} for cat in categories}
for cat in data:
    for p in participants:
        if p in data[cat] and not data[cat][p].empty:
            df = data[cat][p].sort_values('k', ascending=False).reset_index(drop=True)
            switches = df.index[df['Var5'].diff().abs() == 1].tolist()
            switch_pairs = [(switches[i] - 1, switches[i]) for i in range(len(switches)) if switches[i] > 0]
            if len(switch_pairs) == 0:
                rate = np.nan
            elif len(switch_pairs) == 1:
                rate = abs(df.loc[switch_pairs[0][1], 'k'])
            else:
                switch_ks = [abs(df.loc[prev, 'k']) for prev, curr in switch_pairs] + [abs(df.loc[switch_pairs[-1][1], 'k'])]
                rate = gmean([k for k in switch_ks if not np.isnan(k)]) if switch_ks else np.nan
            cumulative_rates[cat][p] = rate
        else:
            cumulative_rates[cat][p] = np.nan

# Step 4: Compute Mean Cumulative Rates and SEM
mean_rates = {}
sem_rates = {}
for cat in categories:
    valid_rates = [r for r in cumulative_rates[cat].values() if not np.isnan(r)]
    mean_rates[cat] = np.mean(valid_rates) if valid_rates else np.nan
    sem_rates[cat] = np.std(valid_rates, ddof=1) / np.sqrt(len(valid_rates)) if valid_rates else np.nan

# Step 5: Create Figure with Subplots
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(10, 5))

# Food Subplot
food_labels = ['Small', 'Large']
food_means = [mean_rates['Small_Food'], mean_rates['Large_Food']]
food_sems = [sem_rates['Small_Food'], sem_rates['Large_Food']]
ax1.bar(food_labels, food_means, yerr=food_sems, capsize=5, color=['skyblue', 'lightcoral'])
ax1.set_title('Food Commodity')
ax1.set_ylabel('Mean Cumulative Rate')
ax1.set_ylim(0, max(food_means) + max(food_sems) * 1.5)

# Money Subplot
money_labels = ['Small', 'Large']
money_means = [mean_rates['Small_Money'], mean_rates['Large_Money']]
money_sems = [sem_rates['Small_Money'], sem_rates['Large_Money']]
ax2.bar(money_labels, money_means, yerr=money_sems, capsize=5, color=['skyblue', 'lightcoral'])
ax2.set_title('Money Commodity')
ax2.set_ylabel('Mean Cumulative Rate')
ax2.set_ylim(0, max(money_means) + max(money_sems) * 1.5)

# Adjust layout and save
plt.tight_layout()
plt.savefig(os.path.join(part2_2_folder, 'part2_2_figure.png'))
plt.close()