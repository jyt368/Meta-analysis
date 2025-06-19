import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
import os
import seaborn as sns

file_path = 'C:/Users/yj368/Meta_analysis/meta_analysis_finall_0527.xlsx'

# Read the Excel file, get all sheet names
xls = pd.ExcelFile(file_path)
sheet_names = xls.sheet_names  

# Iterate over all the sheets (different cognative assessment types) and generate a funnel plot for each assessment type
for sheet in sheet_names:
    df = pd.read_excel(file_path, sheet_name=sheet)

    # Show the head rows to see if the data is loaded correctly
    print(f"Reading data from sheet: {sheet}")
    print(df.head())

    # Extracting values from relevant columns
    try:
        effect_size_column = "Hedges's g"  # take the column named "Hedges's g"
        # the std is next to the effect size column
        std_errors = df.columns[df.columns.get_loc(effect_size_column) + 1] 
        

        study_names = df["Study name"]
        effect_sizes = df[effect_size_column]
        overall_effect = effect_sizes.mean()
        std_errors = df[std_errors]
        print(std_errors)

        # Create the funnel plot for the current subsheet
        plt.figure(figsize=(8, 6))
        plt.scatter(effect_sizes, std_errors, color='black', label='Studies')

        plt.title(f"Funnel Plot for Publication Bias - {sheet}", fontsize=14)
        plt.xlabel("Effect Size (Hedges's g)", fontsize=12)
        plt.ylabel("Standard Error", fontsize=12)

        # Draw the funnel lines (diagonal lines)
        '''z_critical = 1.96  # For 95% confidence interval
        max_std_error = std_errors.max()
        x_range = np.linspace(effect_sizes.min() - 2, effect_sizes.max() + 2, 100)

        # Define the funnel boundary lines (upper and lower) based on standard errors
        upper_funnel = z_critical * std_errors.max() + x_range
        lower_funnel = -z_critical * std_errors.max() + x_range

        # Plot the funnel lines (upper and lower boundaries)
        plt.plot(x_range, upper_funnel, 'k--', label='Upper Funnel Boundary (95% CI)')
        plt.plot(x_range, lower_funnel, 'k--', label='Lower Funnel Boundary (95% CI)')

        # Invert y-axis for better visualization (smaller error at the top)
        plt.gca().invert_yaxis()

        # Set axis limits to ensure the funnel shape is visible
        # plt.xlim(effect_sizes.min() - 0.5, effect_sizes.max() + 0.5)  # Adjust as needed
        plt.ylim(max_std_error * 1.1, 0)  # Ensure the y-axis starts from the maximum standard error

        # Show grid and plot
        plt.grid(False)
        plt.legend()
        plt.tight_layout()'''

        z_critical = 1.96  # for 95% confidence interval
        max_std_error = std_errors.max()
        x_range = np.linspace(effect_sizes.min(), effect_sizes.max(), 100)

        # Define the funnel boundary lines (upper and lower) based on standard errors
        upper_funnel1 = z_critical * std_errors.min() + overall_effect
        upper_funnel2 = z_critical * std_errors.max() + overall_effect
        lower_funnel1 = -z_critical * std_errors.min() + overall_effect
        lower_funnel2 = -z_critical * std_errors.max() + overall_effect


        # Plot the right and left boundaries (horizontal CI lines at the extremes), dash
        plt.plot([upper_funnel1, upper_funnel2], [std_errors.min(), std_errors.max()], color='gray', linestyle='-', lw=1, label="Left Boundary CI")
        plt.plot([lower_funnel1, lower_funnel2], [std_errors.min(), std_errors.max()], color='gray', linestyle='-', lw=1, label="Right Boundary CI")

        # Average effect size in red line
        plt.axvline(overall_effect, color='red', linestyle='--', label='Overall Effect Size')
        plt.gca().invert_yaxis()

        plt.grid(False)
        plt.legend()
        plt.tight_layout()
        

        #save the plot
        output_dir = 'C:/Users/yj368/Meta_analysis/Funnel_plot2/'
        filename = f"funnel_plot_{sheet}.png"  
        sv_path = os.path.join(output_dir, filename)
        plt.savefig(sv_path, dpi=300)  # Save the figure
        print(f"Saved funnel plot for {sheet} as {filename}")

        plt.show()


    except KeyError as e:
        print(f"Missing required columns in sheet {sheet}: {e}")
