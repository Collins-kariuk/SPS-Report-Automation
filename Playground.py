import pandas as pd  # Importing pandas library for data manipulation and analysis
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
import requests  # Importing requests library for making HTTP requests

# Reading the master Excel file containing reports from Zone 1
master_df = pd.read_excel('24 Reports Zone 1.xlsx')
# Reading the 'Activity Report' sheet from the target Excel file
target_df = pd.read_excel('Zone 1 Activity.xlsx', sheet_name='Activity Report')
# Reading the additional Excel file containing induction dates
induction_df = pd.read_excel('MHS Chapters.xlsx')

# Debugging step: Print column names to verify they are as expected
# print("Master DataFrame Columns:", master_df.columns)
# print("Target DataFrame Columns:", target_df.columns)
# print("Induction DataFrame Columns:", induction_df.columns)

# Strip any leading/trailing spaces in column names
target_df.columns = target_df.columns.str.strip()
induction_df.columns = induction_df.columns.str.strip()

def update_induction_date(target_df, induction_df):
    # Loop through each record in the target dataframe
    for index, row in target_df.iterrows():
        school_name = row['Custom Field Data - Chapter School Name']

        # Use fuzzy matching to find the best match for the school name in the induction dataframe
        induction_school_names = induction_df['Institution'].tolist()
        print(f"School Name: {school_name}")
        induction_best_match, induction_score = process.extractOne(school_name, induction_school_names, scorer=fuzz.token_sort_ratio)
        # print the induction best match and induction score
        print(f"Induction Best match: {induction_best_match}, Score: {induction_score}")

        # Check if the best match score is above a certain threshold
        if induction_score > 90:
            # Verify that the match is indeed correct (additional step)
            if school_name == induction_best_match:
                # Step 1: Create a boolean series where the 'Institution' matches the 'induction_best_match'
                matching_rows_boolean_series = induction_df['Institution'] == induction_best_match

                # Step 2: Filter the induction_df to get only the rows where the comparison is True
                matching_rows_df = induction_df[matching_rows_boolean_series]

                # Step 3: Get the index of the matching rows in the filtered DataFrame
                induction_index = matching_rows_df.index

                # If a matching row is found, update the 'Custom Field Data - Last Sigma Pi Sigma Induction Date' field
                if not induction_index.empty:
                    induction_date = induction_df.at[induction_index[0], 'Last Induction']
                    target_df.at[index, 'Custom Field Data - Last Sigma Pi Sigma Induction Date'] = induction_date
    return target_df

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
    # print the score
    # print(f"Best match: {best_match}, Score: {score}")

    # Check if the best match score is above a certain threshold
    if score > 40:
        # Step 1: Select the 'Custom Field Data - Chapter School Name' column from target_df
        school_name_column = target_df['Custom Field Data - Chapter School Name']
        # Step 2: Compare each entry in the school_name_column to the best_match to create a boolean Series
        matching_rows_boolean_series = school_name_column == best_match
        # Step 3: Filter target_df to get only the rows where the comparison is True
        matching_rows_df = target_df[matching_rows_boolean_series]
        # target_index is an Int64Index object containing the indices of the matching rows
        target_index = matching_rows_df.index

        # If a matching row is found, update the relevant fields
        if not target_index.empty:
            # target_index[0] accesses the first element in the Int64Index object
            index = target_index[0]
            # target_df.at[index, 'Custom Field Data - SPS Chapter-Advisor Name'] = master_record['Chapter Adviser Name']
            # target_df.at[index, 'Custom Field Data - SPS Chapter-Advisor E-mail'] = master_record['Chapter Adviser Email']
            target_df.at[index, 'Custom Field Data - SPS Chapter-StudentLeadership-President Name'] = master_record['Incoming SPS President Name']
            target_df.at[index, 'Custom Field Data - SPS Chapter-StudentLeadership-President Email'] = master_record['Incoming SPS President Email']
            # target_df.at[index, 'Custom Field Data - SPS Chapter-StudentLeadership-Vice President Name'] = master_record['Incoming SPS Vice President Name']
            # target_df.at[index, 'Custom Field Data - SPS Chapter-StudentLeadership-Vice President Email'] = master_record['Incoming SPS Vice President Email']
            # target_df.at[index, 'Custom Field Data - SPS Chapter-StudentLeadership-Secretary Name'] = master_record['Incoming SPS Secretary Name']
            # target_df.at[index, 'Custom Field Data - SPS Chapter-StudentLeadership-Secretary Email'] = master_record['Incoming SPS Secretary Email']
            # target_df.at[index, 'Custom Field Data - SPS Chapter-StudentLeadership-Treasurer Name'] = master_record['Incoming SPS Treasurer Name']
            # target_df.at[index, 'Custom Field Data - SPS Chapter-StudentLeadership-Treasurer Email'] = master_record['Incoming SPS Treasurer Email']
            # target_df.at[index, 'Custom Field Data - SPS Chapter-StudentLeadership-Other Officers Names'] = master_record['Other Officers (Format: Name_1; Title_1; Name_2; Title_2 )']
            # target_df.at[index, 'Custom Field Data - SPS Chapter-StudentLeadership-Other Officers Emails'] = master_record['Other Officers Email (Format: email1@mail.edu; email2@mail.edu)']

    # Return the updated target dataframe
    return target_df

# Loop through each record in the master dataframe and update the target dataframe accordingly
for i, row in master_df.iterrows():
    target_df = update_record(row, target_df)

# Update induction dates separately
# target_df = update_induction_date(target_df, induction_df)

# Save the updated target dataframe to a new Excel file
target_df.to_excel('Updated Zone 1 Activity Playground.xlsx', index=False)

# Testing
# school_name = "University of Massachusetts Dartmouth"
# school_names = ["University of Massachusetts - Dartmouth"]
# school_name = "Worcester Polytechnic Institute"
# school_names = ["Worcester Polytechnic Inst"]
# best_match, score = process.extractOne(school_name, school_names, scorer=fuzz.token_sort_ratio)
# print(f"Best match: {best_match}, Score: {score}")
