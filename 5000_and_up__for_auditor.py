import pandas as pd
import numpy as np

filename = input("Enter filename: ")
file_path = filename if filename else "C:/Users/csuar/Downloads/Gifts by Date (1).xlsx"
df = pd.read_excel(file_path, skiprows=3)

# Rename the columns correctly
df.columns = ['Donor ID', 'Name', 'Gift Date', 'Amount', 'Gift Type', 'Reference', 'Solicitation', 'General Ledger']

# Drop rows where 'Name' or 'Amount' is NaN
df = df.dropna(subset=['Name', 'Amount'])

# Clean the 'Amount' column
df['Amount'] = df['Amount'].apply(lambda x: float(str(x).replace('(', '').replace(')', '').replace('$', '').replace(',', '')) if pd.notnull(x) else x)

# Process and aggregate the data
name_to_index = dict()
totals = []
for index, row in df.iterrows():
    name = row['Name']
    amount = row['Amount']
    if name in name_to_index:
        index = name_to_index[name]
        totals[index] = (totals[index][0], totals[index][1] + amount)
    else:
        name_to_index[name] = len(totals)
        totals.append((name, amount))

# Function to extract last name
def extract_last_name(full_name):
    if isinstance(full_name, str):
        parts = full_name.split()
        return parts[-1] if parts else ""
    return ""

# Sort totals by last name
sorted_totals = sorted(totals, key=lambda x: extract_last_name(x[0]))

# Write to file
with open('list1.txt', 'w') as file:
    for name, amount in sorted_totals:
        if amount >= 5000:
            file.write(f"{name} - {amount:.2f}\n")
