import pandas as pd
import numpy as np

filename = input("Enter filename: ")
file_path = filename if filename else "C:/Users/csuar/Downloads/Gifts by Date fy23-24.xlsx"
df = pd.read_excel(file_path, skiprows=2)
df.columns = ['Donor ID', 'Name', 'Gift Date', 'Amount', 'Gift Type', 'Reference', 'Solicitation', 'General Ledger']
df = df.drop(0)

df['Amount'] = df['Amount'].apply(lambda x: float(str(x).replace('(', '').replace(')', '').replace('$', '').replace(',', '')) if pd.notnull(x) and isinstance(x, str) else np.nan)

df = df.dropna(subset=['Amount'])

totals = dict()
for index, row in df.iterrows():
    name = row['Name']
    amount = row['Amount']
    if name in totals:
        totals[name] += amount
    else:
        totals[name] = amount

amount_ranges = [[2500, 4999], [5000, 9999], [10000, 24999], [25000, 49999], [50000, float('inf')]]
sorted_people = [[] for _ in amount_ranges]

def extract_last_name(full_name):
    if isinstance(full_name, str):
        parts = full_name.split()
        return parts[-1] if parts else ""
    return ""

for key, value in totals.items():
    for i, amount_range in enumerate(amount_ranges):
        if amount_range[0] <= value <= amount_range[1]:
            sorted_people[i].append(key)
            break

for i in range(len(sorted_people)):
    sorted_people[i] = sorted(sorted_people[i], key=extract_last_name)

with open('list.txt', 'w') as file:
    for amount_range, people_list in zip(amount_ranges, sorted_people):
        file.write(f"${amount_range[0]}, ${amount_range[1]}\n")
        for person in people_list:
            if isinstance(person, str):
                file.write(f"{person}\n")
        file.write('\n')
