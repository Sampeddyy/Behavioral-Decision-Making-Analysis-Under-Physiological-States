import os
import pandas as pd
import numpy as np
from scipy.stats import wilcoxon, gmean


main_folder = r"S:\lmassignment\LM_A2_2025_data"
part2_1_folder = os.path.join(main_folder, 'part2_1')
os.makedirs(part2_1_folder, exist_ok=True)


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


for cat in data:
    for p in data[cat]:
        df = data[cat][p]
        if not df.empty:
            df['V'] = np.where(df['Var5'] == 0, df['Var1'], (df['Var1'] + df['Var2']) / 2)
            df['k'] = np.where((df['Var4'] > 0) & (df['V'] > 0), (df['Var2'] - df['V']) / (df['V'] * df['Var4']), np.nan)


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

mean_rates = {}
for cat in categories:
    valid_rates = [r for r in cumulative_rates[cat].values() if not np.isnan(r)]
    mean_rates[cat] = np.mean(valid_rates) if valid_rates else np.nan


rel_diff_food = (mean_rates['Large_Food'] - mean_rates['Small_Food']) / mean_rates['Small_Food'] if mean_rates['Small_Food'] != 0 and not np.isnan(mean_rates['Small_Food']) else np.nan
rel_diff_money = (mean_rates['Large_Money'] - mean_rates['Small_Money']) / mean_rates['Small_Money'] if mean_rates['Small_Money'] != 0 and not np.isnan(mean_rates['Small_Money']) else np.nan


food_participants = set(cumulative_rates['Small_Food'].keys()) & set(cumulative_rates['Large_Food'].keys())
small_food_valid = [cumulative_rates['Small_Food'][p] for p in food_participants if not np.isnan(cumulative_rates['Small_Food'][p]) and not np.isnan(cumulative_rates['Large_Food'][p])]
large_food_valid = [cumulative_rates['Large_Food'][p] for p in food_participants if not np.isnan(cumulative_rates['Small_Food'][p]) and not np.isnan(cumulative_rates['Large_Food'][p])]


small_money_valid = [r for r in cumulative_rates['Small_Money'].values() if not np.isnan(r)]
large_money_valid = [r for r in cumulative_rates['Large_Money'].values() if not np.isnan(r)]

stat_food, p_food = wilcoxon(small_food_valid, large_food_valid) if len(small_food_valid) >= 2 else (np.nan, np.nan)
stat_money, p_money = wilcoxon(small_money_valid, large_money_valid) if len(small_money_valid) >= 2 else (np.nan, np.nan)

delay_data = {cat: {} for cat in categories}
for cat in data:
    for p in participants:
        if p in data[cat] and not data[cat][p].empty:
            df = data[cat][p].sort_values('k', ascending=False).reset_index(drop=True)
            switches = df.index[df['Var5'].diff().abs() == 1].tolist()
            delay_data[cat][p] = df.loc[switches[0], 'Var4'] if switches else np.nan
        else:
            delay_data[cat][p] = np.nan

# Food Delays
small_food_delay_valid = [delay_data['Small_Food'][p] for p in food_participants if not np.isnan(delay_data['Small_Food'][p]) and not np.isnan(delay_data['Large_Food'][p])]
large_food_delay_valid = [delay_data['Large_Food'][p] for p in food_participants if not np.isnan(delay_data['Small_Food'][p]) and not np.isnan(delay_data['Large_Food'][p])]

# Money Delays
small_money_delay_valid = [r for r in delay_data['Small_Money'].values() if not np.isnan(r)]
large_money_delay_valid = [r for r in delay_data['Large_Money'].values() if not np.isnan(r)]

stat_delay_food, p_delay_food = wilcoxon(small_food_delay_valid, large_food_delay_valid) if len(small_food_delay_valid) >= 2 else (np.nan, np.nan)
stat_delay_money, p_delay_money = wilcoxon(small_money_delay_valid, large_money_delay_valid) if len(small_money_delay_valid) >= 2 else (np.nan, np.nan)


results_df = pd.DataFrame({
    'Metric': ['Mean Cumulative Rate (Small Food)', 'Mean Cumulative Rate (Large Food)', 
               'Mean Cumulative Rate (Small Money)', 'Mean Cumulative Rate (Large Money)',
               'Relative Rate Difference (Food)', 'Relative Rate Difference (Money)',
               'Wilcoxon Test (Small vs Large Food) Statistic', 'Wilcoxon Test (Small vs Large Food) P-value',
               'Wilcoxon Test (Small vs Large Money) Statistic', 'Wilcoxon Test (Small vs Large Money) P-value',
               'Wilcoxon Test (Delay Small vs Large Food) Statistic', 'Wilcoxon Test (Delay Small vs Large Food) P-value',
               'Wilcoxon Test (Delay Small vs Large Money) Statistic', 'Wilcoxon Test (Delay Small vs Large Money) P-value'],
    'Value': [mean_rates['Small_Food'], mean_rates['Large_Food'], mean_rates['Small_Money'], mean_rates['Large_Money'],
              rel_diff_food, rel_diff_money, stat_food, p_food, stat_money, p_money, 
              stat_delay_food, p_delay_food, stat_delay_money, p_delay_money]
})
results_df.to_excel(os.path.join(part2_1_folder, 'results_part2_1.xlsx'), index=False)


print(f"Mean Cumulative Rate (Small Food): {mean_rates['Small_Food']:.4f}")
print(f"Mean Cumulative Rate (Large Food): {mean_rates['Large_Food']:.4f}")
print(f"Mean Cumulative Rate (Small Money): {mean_rates['Small_Money']:.4f}")
print(f"Mean Cumulative Rate (Large Money): {mean_rates['Large_Money']:.4f}")
print(f"Relative Rate Difference (Food): {rel_diff_food:.4f}")
print(f"Relative Rate Difference (Money): {rel_diff_money:.4f}")
print(f"Wilcoxon Test (Small vs Large Food): Statistic = {stat_food:.4f}, P-value = {p_food:.4f}")
print(f"Wilcoxon Test (Small vs Large Money): Statistic = {stat_money:.4f}, P-value = {p_money:.4f}")
print(f"Wilcoxon Test (Delay Period - Small vs Large Food): Statistic = {stat_delay_food:.4f}, P-value = {p_delay_food:.4f}")
print(f"Wilcoxon Test (Delay Period - Small vs Large Money): Statistic = {stat_delay_money:.4f}, P-value = {p_delay_money:.4f}")