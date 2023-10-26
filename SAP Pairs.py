import pandas as pd
import random

# Load data from the Excel file(name.xlsx)
xls = pd.ExcelFile("SAP HOMIES.xlsx")
#no header, reads the file for seasoned and new members and stores it in varibles
seasoned_members_df = pd.read_excel(xls, "Actives Taking SAPs", header=None, names=["Seasoned Member"])
new_members_df = pd.read_excel(xls, "New Mems!", header=None, names=["New Member"])

# Shuffle the new members and seasoned members randomly
new_members_list = new_members_df["New Member"].tolist()
random.shuffle(new_members_list)

seasoned_members_list = seasoned_members_df["Seasoned Member"].tolist()
random.shuffle(seasoned_members_list)

# Initialize a dictionary to store pairings- basically tracks pairs in key-value style
pairings = {}
# Initialize a dictionary to keep track of the number of pairs for each new member
new_member_pairs = {new_member: 0 for new_member in new_members_list}

# Create pairings for each seasoned member, ensuring each new member has at least 3 pairs and up to 4 if possible
for seasoned_member in seasoned_members_list:
    #Every seasoned member has a corresponding empty group
    pairings[seasoned_member] = []

for _ in range(4):
    #Goes over every seasoned member in the list
    for seasoned_member in seasoned_members_list:
        # Filter new members who have less than 4 pairs and have not been used in this group
        available_new_members = [new_member for new_member in new_members_list if new_member_pairs[new_member] < 4 and new_member not in pairings[seasoned_member]]
        if not available_new_members:
            continue  # No more new members available for this seasoned member so go to next seasoned member
        new_member = random.choice(available_new_members)  # Choose a random new member and add them to the seasoned member's group
        pairings[seasoned_member].append(new_member)
        new_member_pairs[new_member] += 1

# Create a new dataframe(excel sheet) for the pairings
pairings_df = pd.DataFrame(pairings.items(), columns=["Seasoned Member", "New Members"])

# Save the pairings to an Excel file
pairings_df.to_excel("pairingstest.xlsx", index=False)
