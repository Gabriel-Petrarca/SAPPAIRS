import pandas as pd

# Load data from the Excel file (pairings11.xlsx)
pairings_df = pd.read_excel("pairings11.xlsx")

# Initialize an empty list to store all the group lists
all_groups = []

# Iterate through the rows of the DataFrame(excel sheet) to gather all group lists
for _, row in pairings_df.iterrows():
    group = row["New Members"].strip("[]").replace("'", "").split(", ")
    all_groups.append(group)

# Flatten the list of lists (create one big list instead of lots of little lists)
all_names = [name for group in all_groups for name in group]

# Create a Pandas Series from the flattened list (kind of like an excel column)
all_names_series = pd.Series(all_names)

# Count the occurrences of each name
name_counts = all_names_series.value_counts()

# Create a DataFram(excel sheet) from the counts
counts_df = pd.DataFrame({'Name': name_counts.index, 'Count': name_counts.values})

# Save the counts to an Excel file
counts_df.to_excel("all_group_name_counts.xlsx", index=False)
