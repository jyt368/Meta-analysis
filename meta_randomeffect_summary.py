import pandas as pd
import numpy as np
from scipy.stats import norm

input_file = "C:/Users/yj368/Meta_analysis/effect_size_results.xlsx"
df = pd.read_excel(input_file)

# You should have the following columns in your Excel file or rename this when using this script:
df = df.rename(columns={
    "Hedges's g": "g",
    "Lower limit": "lower",
    "Upper limit": "upper"
})

# Calculate SE and variance from CI
z_critical = 1.96  # for 95% confidence interval
df["SE"] = (df["upper"] - df["lower"]) / (2 * z_critical)
df["var"] = df["SE"] ** 2

weights_fixed = 1 / df["var"]
fixed_effect = np.sum(df["g"] * weights_fixed) / np.sum(weights_fixed)
Q = np.sum(weights_fixed * (df["g"] - fixed_effect) ** 2)
df_deg = len(df) - 1
C = np.sum(weights_fixed) - (np.sum(weights_fixed**2) / np.sum(weights_fixed))
tau2 = max(0, (Q - df_deg) / C)

# This is the random-effects model calculation
weights_random = 1 / (df["var"] + tau2)
random_effect = np.sum(df["g"] * weights_random) / np.sum(weights_random)
random_var = 1 / np.sum(weights_random)
random_se = np.sqrt(random_var)
ci_lower = random_effect - z_critical * random_se
ci_upper = random_effect + z_critical * random_se
z_value = random_effect / random_se
p_value = 2 * (1 - norm.cdf(abs(z_value)))
I2 = max(0, (Q - df_deg) / Q) * 100  # percentage

#Create summary output
summary = pd.DataFrame([{
    "Total Hedges' g": round(random_effect, 4),
    "95% CI Lower": round(ci_lower, 4),
    "95% CI Upper": round(ci_upper, 4),
    "Standard Error": round(random_se, 4),
    "Variance": round(random_var, 4),
    "Z-Value": round(z_value, 4),
    "P-Value": round(p_value, 15),#
    "I² (%)": round(I2, 4),
    "Number of Studies": len(df)
}])

#Save summary to Excel
output_file = "C:/Users/yj368/Meta_analysis/meta_summary.xlsx"
summary.to_excel(output_file, index=False)

print(f"Random-effects meta-analysis completed. Results saved to '{output_file}'")
