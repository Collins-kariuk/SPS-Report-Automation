import pandas as pd
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
from typing import List, Tuple

# Reading the master Excel file containing reports from Zone 1
master_df = pd.read_excel('24 Reports Zone 1.xlsx')
# Reading the 'Activity Report' sheet from the target Excel file
target_df = pd.read_excel('Zone 1 Activity.xlsx', sheet_name='Activity Report')
# Reading the additional Excel file containing induction dates
induction_df = pd.read_excel('MHS Chapters.xlsx')

# Strip any leading/trailing spaces in column names
target_df.columns = target_df.columns.str.strip()
induction_df.columns = induction_df.columns.str.strip()

# 'University of Massachusetts - Amherst': 'University of Massachusetts Amherst'
manual_overrides = {
    'Harvard College': 'Harvard University'
}

def get_correct_match(school_name: str, school_names: List[str]) -> Tuple[str, int]:
    if school_name in manual_overrides:
        return manual_overrides[school_name], 100
    else:
        return process.extractOne(school_name, school_names, scorer=fuzz.token_sort_ratio)

# Function to update a record in the target dataframe based on the master dataframe
def update_record(master_record: pd.Series, target_df: pd.DataFrame) -> pd.DataFrame:
    # Extract and trim the school name from the master record
    school_name = master_record['School Name (No abbreviations please)'].strip()
    # print(f"School Name: {school_name}")
    # Use manual overrides or fuzzy matching to find the best match for the school name in the target dataframe
    school_names = [name.strip() for name in target_df['Custom Field Data - Chapter School Name'].tolist()]
    best_match, score = get_correct_match(school_name, school_names)
    # print(f"Best match: {best_match}, Score: {score}")
    # print("-----------------")

    # Check if the best match score is above a certain threshold
    if score > 40:
        # Step 1: Select the 'Custom Field Data - Chapter School Name' column from target_df
        school_name_column = target_df['Custom Field Data - Chapter School Name'].str.strip()
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
            target_df.at[index, 'Custom Field Data - SPS Chapter-Advisor Name'] = master_record['Chapter Adviser Name']
            target_df.at[index, 'Custom Field Data - SPS Chapter-Advisor E-mail'] = master_record['Chapter Adviser Email']
            target_df.at[index, 'Custom Field Data - SPS Chapter-StudentLeadership-President Name'] = master_record['Incoming SPS President Name']
            target_df.at[index, 'Custom Field Data - SPS Chapter-StudentLeadership-President Email'] = master_record['Incoming SPS President Email']
            target_df.at[index, 'Custom Field Data - SPS Chapter-StudentLeadership-Vice President Name'] = master_record['Incoming SPS Vice President Name']
            target_df.at[index, 'Custom Field Data - SPS Chapter-StudentLeadership-Vice President Email'] = master_record['Incoming SPS Vice President Email']
            target_df.at[index, 'Custom Field Data - SPS Chapter-StudentLeadership-Secretary Name'] = master_record['Incoming SPS Secretary Name']
            target_df.at[index, 'Custom Field Data - SPS Chapter-StudentLeadership-Secretary Email'] = master_record['Incoming SPS Secretary Email']
            target_df.at[index, 'Custom Field Data - SPS Chapter-StudentLeadership-Treasurer Name'] = master_record['Incoming SPS Treasurer Name']
            target_df.at[index, 'Custom Field Data - SPS Chapter-StudentLeadership-Treasurer Email'] = master_record['Incoming SPS Treasurer Email']
            target_df.at[index, 'Custom Field Data - SPS Chapter-StudentLeadership-Other Officers Names'] = master_record['Other Officers (Format: Name_1; Title_1; Name_2; Title_2 )']
            target_df.at[index, 'Custom Field Data - SPS Chapter-StudentLeadership-Other Officers Emails'] = master_record['Other Officers Email (Format: email1@mail.edu; email2@mail.edu)']

    # Return the updated target dataframe
    return target_df

# Function to update the chapter reports based on the master dataframe and current year
def update_chapter_reports(target_df: pd.DataFrame, master_df: pd.DataFrame, current_year: int) -> pd.DataFrame:
    for index, row in master_df.iterrows():
        school_name = row['School Name (No abbreviations please)'].strip()
        print(f"School Name: {school_name}")
        school_names = [name.strip() for name in target_df['Custom Field Data - Chapter School Name'].tolist()]
        best_match, score = get_correct_match(school_name, school_names)
        print(f"Best match: {best_match}, Score: {score}")
        print("-----------------")

        if score > 40:
            boolean_series = target_df['Custom Field Data - Chapter School Name'] == best_match
            selected_rows = target_df.loc[boolean_series, 'Custom Field Data - Chapter Reports']
            values_array = selected_rows.values
            current_entry = values_array[0]

            current_entry_str = str(current_entry) if pd.notna(current_entry) else ''

            if pd.isna(current_entry):
                updated_entry = str(current_year)
            else:
                if str(current_year) not in current_entry_str:
                    updated_entry = f'{current_entry_str}; {current_year}'
                else:
                    updated_entry = current_entry_str

            target_df.loc[boolean_series, 'Custom Field Data - Chapter Reports'] = updated_entry

    return target_df

# Loop through each record in the master dataframe and update the target dataframe accordingly
# for i, row in master_df.iterrows():
#     target_df = update_record(row, target_df)

# Update chapter reports based on master_df and current year
target_df = update_chapter_reports(target_df, master_df, 2024)

# Save the updated target dataframe to a new Excel file
target_df.to_excel('Updated Zone 1 Activity Playground.xlsx', index=False)

# Testing
# school_name = "Harvard College"
# school_names = ["Harvard University", "Marlboro College"]
# best_match, score = process.extractOne(school_name, school_names, scorer=fuzz.token_sort_ratio)
# print(f"Best match: {best_match}, Score: {score}")
