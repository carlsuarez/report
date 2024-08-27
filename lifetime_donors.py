import pandas as pd
import numpy as np
import re

filename = input("Enter filename: ")
file_path = filename if filename else "C:/Users/csuar/Downloads/Lifetime Giving (3).xlsx"
df = pd.read_excel(file_path, skiprows=1)
df.columns = ['First Name', 'Last Name', 'Spouse/Partner Full Name', 'Date of Gift', 'Gift Amount', 'General Ledger']
df = df.drop(0)

# Function to clean and convert the 'Gift Amount'
def clean_amount(amount):
    if pd.notnull(amount) and isinstance(amount, str):
        cleaned_amount = re.sub(r'[()\$,]', '', amount)
        try:
            return float(cleaned_amount)
        except ValueError:
            return np.nan
    return np.nan

df['Gift Amount'] = df['Gift Amount'].apply(clean_amount)

# Dictionary to store donors and their total gift amounts
donors = {}

for index, row in df.iterrows():
    first_name = row['First Name'] if pd.notna(row['First Name']) else ''
    last_name = row['Last Name'] if pd.notna(row['Last Name']) else ''
    spouse = row['Spouse/Partner Full Name'] if pd.notna(row['Spouse/Partner Full Name']) else ''
    amount = row['Gift Amount']
    
    if pd.isna(amount) or first_name == 'Anonymous/no name provided' or last_name == 'Anonymous/no name provided':
        continue

    donor_key = (first_name, last_name, spouse)
    
    if donor_key in donors:
        donors[donor_key] += amount
    else:
        donors[donor_key] = amount

amount_ranges = [[float('-inf'), 4999], [5e3, 9999], [10e3, 24999], [25e3, 49999], [50e3, 99999], [100e3, 499999], [500e3, 999999], [1e6, float('inf')]]
sorted_people = [[] for _ in amount_ranges]

for key, value in donors.items():
    for i, amount_range in enumerate(amount_ranges):
        if amount_range[0] <= value <= amount_range[1]:
            sorted_people[i].append(key)
            break

for i in range(len(sorted_people)):
    sorted_people[i] = sorted(sorted_people[i], key=lambda x: x[1])

with open('list2.txt', 'w') as file:
    for amount_range, people_list in zip(amount_ranges, sorted_people):
        file.write(f"${amount_range[0]:,.0f} - ${amount_range[1]:,.0f}\n")
        for fn, ln, sp in people_list:
            if not fn: # Case where there is no first name (company or organization)
                file.write(f"{ln}\n")
            elif not sp: # Case where there is no spouse name
                file.write(f"{fn} {ln}\n")
            else: # Case where the first name, last name and spouse name are defined
                file.write(f"{sp} and {fn} {ln}\n")
        file.write('\n')
