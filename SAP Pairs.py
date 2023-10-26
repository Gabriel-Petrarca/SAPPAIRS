import pandas as pd
import random

# Load data from the Excel file
xls = pd.ExcelFile("SAP HOMIES.xlsx")
seasoned_members_df = pd.read_excel(xls, "Actives Taking SAPs")
new_members_df = pd.read_excel(xls, "New Mems!")

# Shuffle the new members randomly
new_members_list = new_members_df["New Member"].tolist()
random.shuffle(new_members_list)

# Initialize a dictionary to store pairings
pairings = {}

# Create pairings for each seasoned member
for _, seasoned_member in seasoned_members_df.iterrows():
    seasoned_member_name = seasoned_member["Seasoned Member"]
    pairings[seasoned_member_name] = []

    # Get three different new members
    for _ in range(3):
        new_member = new_members_list.pop(0)  # Remove the new member from the list
        pairings[seasoned_member_name].append(new_member)

# Create a new dataframe for the pairings
pairings_df = pd.DataFrame(pairings.items(), columns=["Seasoned Member", "New Members"])

# Save the pairings to an Excel file
pairings_df.to_excel("pairings.xlsx", index=False)
