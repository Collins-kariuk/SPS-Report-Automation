import pandas as pd  # Importing pandas library for data manipulation and analysis
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
import requests  # Importing requests library for making HTTP requests

# Reading the master Excel file containing reports from Zone 1
master_df = pd.read_excel('24 Reports Zone 1.xlsx')

# Reading the 'Activity Report' sheet from the target Excel file
target_df = pd.read_excel('Zone 1 Activity.xlsx', sheet_name='Activity Report')

# Debugging step: Print column names to verify they are as expected
print("Master DataFrame Columns:", master_df.columns)
print("Target DataFrame Columns:", target_df.columns)

# Strip any leading/trailing spaces in column names
target_df.columns = target_df.columns.str.strip()

# Function to update a record in the target dataframe based on the master dataframe
def update_record(master_record, target_df):
    # Extract the school name from the master record
    school_name = master_record['School Name (No abbreviations please)']

    # Use fuzzy matching to find the best match for the school name in the target dataframe
    school_names = target_df['Custom Field Data - Chapter School Name'].tolist()

    # Arguments:
    # - `school_name`: The string you want to find a match for.
    # - `school_names`: The list of potential matches.
    # - `scorer=fuzz.token_sort_ratio`: Specifies the scoring function used to evaluate the
    # similarity between strings. `fuzz.token_sort_ratio` is a function that compares strings by
    # sorting the tokens (words) in each string and then computing a ratio of similarity.

    # Returns:
    # - `best_match`: The string from `school_names` that has the highest similarity score to
    # `school_name`.
    # - `score`: The similarity score between `school_name` and `best_match`. This score ranges
    # from 0 to 100, where 100 means an exact match.
    best_match, score = process.extractOne(school_name, school_names, scorer=fuzz.token_sort_ratio)

    # Check if the best match score is above a certain threshold (e.g., 90)
    if score > 40:
        target_index = target_df[target_df['Custom Field Data - Chapter School Name'] == best_match].index
        # If a matching row is found, update the relevant fields
        if not target_index.empty:
            index = target_index[0]
            target_df.at[index, 'Custom Field Data - SPS Chapter-Advisor Name'] = master_record['Chapter Adviser Name']
            target_df.at[index, 'Custom Field Data - SPS Chapter-Advisor E-mail'] = master_record['Chapter Adviser Email']

    # Return the updated target dataframe
    return target_df

# Loop through each record in the master dataframe and update the target dataframe accordingly
for i, row in master_df.iterrows():
    target_df = update_record(row, target_df)

# Save the updated target dataframe to a new Excel file
target_df.to_excel('Updated Zone 1 Activity New.xlsx', index=False)

# Note: Add any additional processing or saving of the target_df if required

## Testing
# school_name = "Harvard College"
# school_names = ["Harvard University"]
# # school_name = "University of Massachusetts - Amherst"
# # school_names = ["University of Massachusetts Amherst"]
# best_match, score = process.extractOne(school_name, school_names, scorer=fuzz.token_sort_ratio)
# print(f"Best match: {best_match}, Score: {score}")