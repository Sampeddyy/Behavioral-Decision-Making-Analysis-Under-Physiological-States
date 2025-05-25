import os
import pandas as pd
import numpy as np
from scipy.stats import wilcoxon, gmean

# Main folder
main_folder = r"S:\lmassignment\LM_A2_2025_data"
part2_3_folder = os.path.join(main_folder, 'part2_3')
os.makedirs(part2_3_folder, exist_ok=True)

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

# Step 3: Compute Cumulative Rates and Delay Data
cumulative_rates = {cat: {} for cat in categories}
delay_data = {cat: {} for cat in categories}
for cat in data:
    for p in participants:
        if p in data[cat] and not data[cat][p].empty:
            df = data[cat][p].sort_values('k', ascending=False).reset_index(drop=True)
            switches = df.index[df['Var5'].diff().abs() == 1].tolist()
            switch_pairs = [(switches[i] - 1, switches[i]) for i in range(len(switches)) if switches[i] > 0]
            if len(switch_pairs) == 0:
                rate = np.nan
                delay = np.nan
            elif len(switch_pairs) == 1:
                rate = abs(df.loc[switch_pairs[0][1], 'k'])
                delay = df.loc[switch_pairs[0][1], 'Var4']
            else:
                switch_ks = [abs(df.loc[prev, 'k']) for prev, curr in switch_pairs] + [abs(df.loc[switch_pairs[-1][1], 'k'])]
                rate = gmean([k for k in switch_ks if not np.isnan(k)]) if switch_ks else np.nan
                delay = df.loc[switches[0], 'Var4']  # First switch delay
            cumulative_rates[cat][p] = rate
            delay_data[cat][p] = delay
        else:
            cumulative_rates[cat][p] = np.nan
            delay_data[cat][p] = np.nan

# Step 4: Perform Wilcoxon Tests
# Food: Cumulative Rates
food_participants = set(cumulative_rates['Small_Food'].keys()) & set(cumulative_rates['Large_Food'].keys())
small_food_rates = [cumulative_rates['Small_Food'][p] for p in food_participants if not np.isnan(cumulative_rates['Small_Food'][p]) and not np.isnan(cumulative_rates['Large_Food'][p])]
large_food_rates = [cumulative_rates['Large_Food'][p] for p in food_participants if not np.isnan(cumulative_rates['Small_Food'][p]) and not np.isnan(cumulative_rates['Large_Food'][p])]
differences = np.array(small_food_rates) - np.array(large_food_rates)
ranks = np.argsort(np.abs(differences)) + 1
positive_ranks = ranks[differences > 0]
stat_food_rates = np.sum(positive_ranks)  # Sum of positive ranks (153 if all positive)
p_food_rates = wilcoxon(small_food_rates, large_food_rates)[1] if len(small_food_rates) >= 2 else np.nan

# Food: Delay Periods
small_food_delays = [delay_data['Small_Food'][p] for p in food_participants if not np.isnan(delay_data['Small_Food'][p]) and not np.isnan(delay_data['Large_Food'][p])]
large_food_delays = [delay_data['Large_Food'][p] for p in food_participants if not np.isnan(delay_data['Small_Food'][p]) and not np.isnan(delay_data['Large_Food'][p])]
stat_food_delays, p_food_delays = wilcoxon(small_food_delays, large_food_delays) if len(small_food_delays) >= 2 else (np.nan, np.nan)

# Money: Cumulative Rates
money_participants = set(cumulative_rates['Small_Money'].keys()) & set(cumulative_rates['Large_Money'].keys())
small_money_rates = [cumulative_rates['Small_Money'][p] for p in money_participants if not np.isnan(cumulative_rates['Small_Money'][p]) and not np.isnan(cumulative_rates['Large_Money'][p])]
large_money_rates = [cumulative_rates['Large_Money'][p] for p in money_participants if not np.isnan(cumulative_rates['Small_Money'][p]) and not np.isnan(cumulative_rates['Large_Money'][p])]
stat_money_rates, p_money_rates = wilcoxon(small_money_rates, large_money_rates) if len(small_money_rates) >= 2 else (np.nan, np.nan)

# Money: Delay Periods
small_money_delays = [delay_data['Small_Money'][p] for p in money_participants if not np.isnan(delay_data['Small_Money'][p]) and not np.isnan(delay_data['Large_Money'][p])]
large_money_delays = [delay_data['Large_Money'][p] for p in money_participants if not np.isnan(delay_data['Small_Money'][p]) and not np.isnan(delay_data['Large_Money'][p])]
stat_money_delays, p_money_delays = wilcoxon(small_money_delays, large_money_delays) if len(small_money_delays) >= 2 else (np.nan, np.nan)

# Step 5: Save Results
results_df = pd.DataFrame({
    'Metric': [
        'Wilcoxon Test (Small vs Large Food Rates) Statistic', 'Wilcoxon Test (Small vs Large Food Rates) P-value',
        'Wilcoxon Test (Small vs Large Food Delays) Statistic', 'Wilcoxon Test (Small vs Large Food Delays) P-value',
        'Wilcoxon Test (Small vs Large Money Rates) Statistic', 'Wilcoxon Test (Small vs Large Money Rates) P-value',
        'Wilcoxon Test (Small vs Large Money Delays) Statistic', 'Wilcoxon Test (Small vs Large Money Delays) P-value'
    ],
    'Value': [
        stat_food_rates, p_food_rates,
        stat_food_delays, p_food_delays,
        stat_money_rates, p_money_rates,
        stat_money_delays, p_money_delays
    ]
})
results_df.to_excel(os.path.join(part2_3_folder, 'results_part2_3.xlsx'), index=False)

# Print Only Required Outputs with Higher Precision
print(f"Wilcoxon Test (Small vs Large Food Rates): Statistic = {stat_food_rates:.2f}, P-value = {p_food_rates:.6f}")
print(f"Wilcoxon Test (Small vs Large Food Delays): Statistic = {stat_food_delays:.2f}, P-value = {p_food_delays:.2f}")
print(f"Wilcoxon Test (Small vs Large Money Rates): Statistic = {stat_money_rates:.2f}, P-value = {p_money_rates:.6f}")
print(f"Wilcoxon Test (Small vs Large Money Delays): Statistic = {stat_money_delays:.2f}, P-value = {p_money_delays:.2f}")