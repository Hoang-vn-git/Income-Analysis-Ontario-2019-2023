import pandas as pd

# Read the CSV file
df = pd.read_csv("11100239.csv")

# Remove duplicate rows to avoid counting the same data multiple times
df = df.drop_duplicates()

# Rename columns for easier use
df = df.rename(columns={
    "REF_DATE": "Year",   # Year of the data
    "GEO": "Region",      # Geographic region
    "VALUE": "Income"     # Income value
})

# Replace missing or special values with 0
df['Income'] = df['Income'].fillna(0)

# Convert data types to ensure proper numeric operations
df["Year"] = df["Year"].astype(int)
df["Income"] = df["Income"].astype(float)
df['Age group'] = df['Age group'].astype(str)
df['Sex'] = df['Sex'].astype(str)

# Filter data:
# - exclude rows with "Both sexes" to separate male/female analysis
# - exclude "Total income" rows to focus on specific income sources
# - keep only the Ontario region
# - only include data from 2019 onwards
# - exclude aggregated age groups for more detailed analysis
df = df[(df["Sex"] != "Both sexes") &
        (df["Region"] == "Ontario") &
        (df["Year"] >= 2019) &
        (df["Age group"] != "15 years and over") &
        (df["Age group"] != "25 to 54 years")
]

# Keep only the necessary columns for analysis
columns_needed = ["Year","Region","Age group","Sex","Income source","Statistics",'SCALAR_FACTOR',"Income"]
df = df[columns_needed]

# Separate data into Average income and Median income
avg_df = df[df["Statistics"].str.contains("Average income")]
med_df = df[df["Statistics"].str.contains("Median income")]

# Pivot tables to reshape data:
# Rows: Year, Sex, Age group
# Columns: Income source
# Values: Income
avg_pivot = avg_df.pivot_table(index=["Year", "Sex", "Age group"], columns="Income source", values="Income").reset_index()
med_pivot = med_df.pivot_table(index=["Year", "Sex", "Age group"], columns="Income source", values="Income").reset_index()

# Calculate yearly trends (mean for average income, median for median income)
avg_trend = avg_pivot.groupby("Year").mean(numeric_only=True).reset_index()
med_trend = med_pivot.groupby("Year").median(numeric_only=True).reset_index()

# Gender comparison: average income by year and sex
avg_gender = avg_pivot.groupby(['Year','Sex']).mean(numeric_only=True).reset_index()
med_gender = med_pivot.groupby(['Year','Sex']).mean(numeric_only=True).reset_index()

# Age group comparison: average income by year and age group
avg_age = avg_pivot.groupby(['Year','Age group']).mean(numeric_only=True).reset_index()
med_age = med_pivot.groupby(['Year','Age group']).mean(numeric_only=True).reset_index()

# Select income sources related to government transfers
gov_sources = ['COVID-19 benefits', 'Canada Pension Plan (CPP) and Quebec Pension Plan (QPP) benefits',
               'Child benefits', 'Employment Insurance (EI) benefits', 'Government transfers',
               'Old Age Security (OAS) and Guaranteed Income Supplement (GIS)', 'Other government transfers',
               'Social assistance']

# Calculate government transfers trend by year
gov_trend = avg_pivot.groupby('Year')[gov_sources].sum().reset_index()

# Export pivot tables to Excel with separate sheets for Average and Median income
with pd.ExcelWriter("income_analysis.xlsx", engine='xlsxwriter') as writer:
    # General
    avg_pivot.to_excel(writer, sheet_name="Average Income")
    med_pivot.to_excel(writer, sheet_name="Median Income")
    # Yearly trend
    avg_trend.to_excel(writer, sheet_name="Avg Trend", index=False)
    med_trend.to_excel(writer, sheet_name="Med Trend", index=False)

    # Gender comparison
    avg_gender.to_excel(writer, sheet_name="Avg Gender", index=False)
    med_gender.to_excel(writer, sheet_name="Med Gender", index=False)

    # Age group comparison
    avg_age.to_excel(writer, sheet_name="Avg Age", index=False)
    med_age.to_excel(writer, sheet_name="Med Age", index=False)

    # Government transfers
    gov_trend.to_excel(writer, sheet_name="Gov Transfers", index=False)