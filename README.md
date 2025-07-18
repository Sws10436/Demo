# Demo
College Generator
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import time


# Load the Excel file (adjust path if needed)
df = pd.read_excel("C:/Users/Santosh/Desktop/My Folder/Josaa.xlsx", sheet_name="Ranks")

# Drop rows where Place is missing (optional)
df = df[df["Place"].notna()]

# Clean spaces and get unique values
unique_places = df["Place"].str.strip().unique()

#Defining required functions
def type_slow(input_str):
    for s in str(input_str):
        print(s, end="")
        time.sleep(0.02)
    time.sleep(len(input_str)*0.02)
    print()




# Input
RANK = int(input("Enter your AIR: "))
Gender=int(input("Enter your gender,1)Gender-Neutral,2)Female-only(including Supernumerary):"))
type_slow("Enter the index number of the place you want to specificly find the branch for or just leave it blank to view all possibilities")
for i, place in enumerate(sorted(unique_places), start=1):
    type_slow(f"{i}. {place}")
choice_input= input("Enter index number (or leave blank):").strip()
unique_places_df = pd.DataFrame(sorted(unique_places), columns=["Place"])

filter_by_place = False
place = ""

if choice_input:
    try:
        choice = int(choice_input)
        if 1 <= choice <= len(unique_places_df):
            place = unique_places_df.iloc[choice - 1]["Place"]
            filter_by_place = True
        else:
            print("⚠️ Invalid index. Showing all places.")
    except ValueError:
        print("⚠️ Invalid input. Showing all places.")

if Gender==1:
    gender="Gender-Neutral"
elif Gender==2:
    gender="Female-only (including Supernumerary)"
else:
    print("invalid input")
# Load workbook and worksheet
wb = load_workbook('C:/Users/Santosh/Desktop/My Folder/Josaa.xlsx')
ws = wb['Ranks']

# Collect matching rows
matching_rows = []

for row in range(2, 555):
    # Get closing rank
    rank_char = get_column_letter(6)
    cell_value = ws[rank_char + str(row)].value

    if cell_value is None:
        continue

    try:
        closing_rank = int(cell_value)
    except ValueError:
        continue

    # Check if within rank range
    if RANK <= closing_rank:
        # Check place
        place_char = get_column_letter(2)
        place_value = ws[place_char + str(row)].value

        gender_char = get_column_letter(4)
        if gender==ws[gender_char + str(row)].value:
            if not filter_by_place or place.lower() == str(place_value).strip().lower():
                row_data = []
                for col in range(1, 7):
                    char = get_column_letter(col)
                    row_data.append(ws[char + str(row)].value)
                    matching_rows.append(row_data)

        

# Display results as a table
if matching_rows:
    df = pd.DataFrame(matching_rows, columns=["Institute", "Place", "Branch", "Gender", "Opening Rank", "Closing Rank"])
    print("\nAvailable Options Based on Your Rank and Preferred Location:\n")
    print(df.to_string(index=False))
else:
    print("\nNo Branches  found for the given AIR and place.")

print("Thanks for using this Branch Generating Tool! We hope you are satisfies with the information we provide.")
print("If we werent able to do so we are sorry ")
print()
