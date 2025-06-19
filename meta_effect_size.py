import pandas as pd
import math
from scipy.stats import norm

# Compute effect sizes with given data of means, std, and sample sizes from each group
def compute_effect_sizes(row):
    mean1 = row["Treated mean"]
    mean2 = row["Control mean"]
    sd1 = row["Treated std"]
    sd2 = row["Control std"]
    n1 = row["Treated number"]
    n2 = row["Control number"]
    direction = str(row["Effect direction"]).strip().lower() if isinstance(row["Effect direction"], str) else None

    # Effect direction
    diff_means = mean1 - mean2
    if direction == "negative":
        diff_means = -abs(diff_means)
    else:  # Assume "positive" by default
        diff_means = abs(diff_means)

    # Pooled standard deviation
    spooled = math.sqrt(((sd1**2) * (n1 - 1) + (sd2**2) * (n2 - 1)) / (n1 + n2 - 2))

    # Cohen's d
    d = diff_means / spooled

    # Variance and standard error of d
    vd = (n1 + n2) / (n1 * n2) + (d**2 / (2 * (n1 + n2)))
    sed = math.sqrt(vd)

    # Hedges’ correction
    j = 1 - (3 / (4 * (n1 + n2 - 2) - 1))
    g = d * j
    vg = vd * j**2
    seg = math.sqrt(vg)

    # CI for Hedges’ g
    ci_lower = g - 1.96 * seg
    ci_upper = g + 1.96 * seg

    # Z-value and p-value
    z_value = g / seg
    p_value = 2 * (1 - norm.cdf(abs(z_value)))

    # Mean difference variance and SE
    var_diff_means = (sd1**2 / n1) + (sd2**2 / n2)
    se_diff_means = math.sqrt(var_diff_means)

    return pd.Series([
        diff_means,
        se_diff_means,
        var_diff_means,
        d,
        sed,
        vd,
        g,
        seg,
        vg,
        ci_lower,
        ci_upper,
        z_value,
        p_value
    ], index=[
        "Difference in means",
        "SE (diff)",
        "Var (diff)",
        "Cohen's d",
        "SE (d)",
        "Var (d)",
        "Hedges' g",
        "SE (g)",
        "Var (g)",
        "Hedges' g CI Lower",
        "Hedges' g CI Upper",
        "Z-Value",
        "p-Value"
    ])

# Give the file with means, std, and sample sizes from each group from the literature 
input_file = "C:/Users/yj368/Meta_analysis/effect1.xlsx" 
df = pd.read_excel(input_file)
#print(df[["Treated mean", "Control mean", "Treated std", "Control std", "Treated number", "Control number"]].head())

results = df.apply(compute_effect_sizes, axis=1)

# Include both original data and computed effect sizes in the final file
final_df = pd.concat([df, results], axis=1)

# Save
output_file = "C:/Users/yj368/Meta_analysis/effect_size_results.xlsx"
final_df.to_excel(output_file, index=False)

print(f"Effect sizes calculated and saved to '{output_file}'")
