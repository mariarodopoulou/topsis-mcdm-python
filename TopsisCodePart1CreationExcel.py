"""
Interactive TOPSIS Decision Analysis System
Author: Maria Rodopoulou

Description:
This system allows users to define custom criteria,
alternatives, weights and criterion types, then
automatically generates an Excel input file for
TOPSIS multi-criteria decision analysis.

Technologies:
- Python
- Pandas
- OpenPyXL
"""
import pandas as pd
from openpyxl import Workbook

def get_integer_input(prompt):
    while True:
        try:
            return int(input(prompt))
        except ValueError:
            print("Please enter a valid integer.")

def get_float_list(prompt):
    while True:
        try:
            return list(map(float, input(prompt).split()))
        except ValueError:
            print("Please enter valid numeric values separated by spaces.")

def get_criteria_types(prompt, count):
    while True:
        try:
            values = list(map(int, input(prompt).split()))
            if len(values) == count and all(v in [0, 1] for v in values):
                return ["Benefit" if v == 1 else "Cost" for v in values]
            else:
                raise ValueError
        except ValueError:
            print("Please enter 0 or 1 for each criterion separated by spaces.")

# load number of criteria and alternatives
num_criteria = get_integer_input("Enter number of criteria: ")
num_alternatives = get_integer_input("Enter number of alternatives: ")

# Name insert
criteria_names = [input(f"Criterion name {i+1}: ") for i in range(num_criteria)]
alternative_names = [input(f"Alternative name {i+1}: ") for i in range(num_alternatives)]

# Create Decision Matrix
decision_matrix = []
for name in alternative_names:
    decision_matrix.append(get_float_list(f"Values for {name} (space separated): "))

decision_df = pd.DataFrame(decision_matrix, columns=criteria_names, index=alternative_names)

# Download weights
weights = get_float_list("\nEnter weights for each criterion (space separated): ")
weights_df = pd.DataFrame({'Criterion': criteria_names, 'Weight': weights})

# Categorization of criteria (benefit or cost criteria)
criteria_types = get_criteria_types("\nEnter 1 for benefit, 0 for cost criteria (space separated): ", num_criteria)
criteria_types_df = pd.DataFrame({'Criterion': criteria_names, 'Type': criteria_types})

# Create and save file Excel
file_name = "topsis_input.xlsx"
with pd.ExcelWriter(file_name) as writer:
    decision_df.to_excel(writer, sheet_name="Decision_Matrix", index=True, index_label="Alternative")
    weights_df.to_excel(writer, sheet_name="Criteria_Weights", index=False)
    criteria_types_df.to_excel(writer, sheet_name="Criteria_Types", index=False)

print(f"\nFile '{file_name}' created successfully!")