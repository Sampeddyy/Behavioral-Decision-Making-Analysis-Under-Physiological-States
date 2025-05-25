import os
import pandas as pd
import numpy as np
from scipy.stats import wilcoxon

# Main folder
main_folder = r"S:\lmassignment\LM_A2_2025_data"
part6_folder = os.path.join(main_folder, 'part 6')

# Create part 6 folder
os.makedirs(part6_folder, exist_ok=True)

# Load individual cumulative rates from part 3
participants = [f for f in os.listdir(main_folder) if os.path.isdir(os.path.join(main_folder, f))]
categories = ['food-C', 'food-F', 'money-C', 'money-F']
rates = {cat: [] for cat in categories}

for p in participants:
    part3_folder = os.path.join(main_folder, p, 'part 3')
    for cat in categories:
        file = os.path.join(part3_folder, f"{cat}_cumulative_rate.xlsx")
        try:
            if os.path.exists(file):
                df = pd.read_excel(file)
                rate = df['Cumulative_Rate'].iloc[0]
                rates[cat].append(rate if not pd.isna(rate) else np.nan)
            else:
                rates[cat].append(np.nan)
        except Exception as e:
            print(f"Error reading {file}: {e}")
            rates[cat].append(np.nan)

# Compute relative differences per participant
food_diffs = []
money_diffs = []
for i in range(len(participants)):
    food_c, food_f = rates['food-C'][i], rates['food-F'][i]
    money_c, money_f = rates['money-C'][i], rates['money-F'][i]
    food_diff = (food_f - food_c) / food_c if not np.isnan(food_c) and not np.isnan(food_f) and food_c != 0 else np.nan
    money_diff = (money_f - money_c) / money_c if not np.isnan(money_c) and not np.isnan(money_f) and money_c != 0 else np.nan
    food_diffs.append(food_diff)
    money_diffs.append(money_diff)

# Pairwise removal of NaN for Wilcoxon test
test_df = pd.DataFrame({'Food_Relative_Difference': food_diffs, 'Money_Relative_Difference': money_diffs}).dropna()
print(f"Number of valid participant pairs: {len(test_df)}")

# Perform Wilcoxon Signed-Rank Test
if len(test_df) < 2:
    print("Error: Not enough valid pairs for Wilcoxon test (need at least 2).")
else:
    try:
        stat, p_value = wilcoxon(test_df['Food_Relative_Difference'], test_df['Money_Relative_Difference'])
        print(f"Wilcoxon Signed-Rank Test Results:")
        print(f"Test Statistic: {stat:.4f}")
        print(f"P-value: {p_value:.4f}")

        # Save test results to part 6
        results_df = pd.DataFrame({'Test': ['Wilcoxon Signed-Rank'], 'Statistic': [stat], 'P-value': [p_value]})
        results_df.to_excel(os.path.join(part6_folder, 'wilcoxon_test_results.xlsx'), index=False)
        print(f"Results saved to {part6_folder}/wilcoxon_test_results.xlsx")
    except Exception as e:
        print(f"Error performing Wilcoxon test: {e}")

print("Question 1(vi) processing complete!")