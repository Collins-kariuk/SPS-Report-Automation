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

def update_induction_date(school_name, city, target_df, index, induction_df):
    # Combine school name and city for matching
    school_city_combined = school_name + " " + city
    induction_combined = induction_df['Institution'] + " " + induction_df['City']

    # Use fuzzy matching to find the best match for the combined school name and city
    result = process.extractOne(school_city_combined, induction_combined.tolist(), scorer=fuzz.token_sort_ratio)

    # Ensure that result is not None and contains exactly two values
    if result and len(result) == 2:
        induction_best_match, induction_score = result
    else:
        induction_best_match, induction_score = None, 0

    # Check if the best match score is above a certain threshold
    if induction_score > 90:
        # Step 1: Create a boolean series where the combined field matches the best match
        matching_rows_boolean_series = (induction_df['Institution'] + " " + induction_df['City']) == induction_best_match

        # Step 2: Filter the induction_df to get only the rows where the comparison is True
        matching_rows_df = induction_df[matching_rows_boolean_series]

        # Step 3: Get the index of the matching rows in the filtered DataFrame
        induction_index = matching_rows_df.index

        # If a matching row is found, update the 'Custom Field Data - Last Sigma Pi Sigma Induction Date' field
        if not induction_index.empty:
            induction_date = induction_df.at[induction_index[0], 'Last Induction']
            target_df.at[index, 'Custom Field Data - Last Sigma Pi Sigma Induction Date'] = induction_date
    
    return target_df

# Function to update a record in the target dataframe based on the master dataframe and induction dataframe
def update_record(master_record, target_df, induction_df):
    # Extract the school name from the master record
    school_name = master_record['School Name (No abbreviations please)']
    # Step 1: Create a Boolean Series to match the school name
    boolean_series = target_df['Custom Field Data - Chapter School Name'] == school_name
    # Step 2: Filter the DataFrame using the Boolean Series
    filtered_df = target_df[boolean_series]

    # # Check if there are any matching rows
    if filtered_df.empty:
        print(f"No matching school name found for {school_name}")
        return target_df

    # Step 3: Get the index of the first matching row
    matching_index = filtered_df.index[0]
    # Step 4: Check if the 'City' column exists and access the value
    if 'Member/Non-Member - Employer City' not in target_df.columns:
        print(f"City column not found for {school_name}")
        city = "Default City"  # Use default city value
    else:
        city = target_df.at[matching_index, 'Member/Non-Member - Employer City']

    # Check if the city value is valid (not null or empty)
    if pd.isna(city) or city == '':
        print(f"No city information available for {school_name}, using default city")
        city = "Default City" # Use default city value

    # Combine school name and city for matching
    school_city_combined = school_name + " " + city
    target_combined = target_df['Custom Field Data - Chapter School Name'] + " " + target_df['Member/Non-Member - Employer City']
    # print(f"Matching: {school_city_combined} against {target_combined.tolist()}")

    # Use fuzzy matching to find the best match for the combined school name and city in the target dataframe
    result = process.extractOne(school_city_combined, target_combined.tolist(), scorer=fuzz.token_sort_ratio)

    # Ensure that result is not None and contains exactly two values
    if result and len(result) == 2:
        best_match, score = result
    else:
        best_match, score = None, 0

    # Check if the best match score is above a certain threshold
    if score > 40:
        # Step 1: Select the combined 'Custom Field Data - Chapter School Name' and 'City' column from target_df
        combined_column = target_df['Custom Field Data - Chapter School Name'] + " " + target_df['Member/Non-Member - Employer City']
        # Step 2: Compare each entry in the combined_column to the best_match to create a boolean Series
        matching_rows_boolean_series = combined_column == best_match
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
            # target_df.at[index, 'Custom Field Data - SPS Chapter-StudentLeadership-President Name'] = master_record['Incoming SPS President Name']
            # target_df.at[index, 'Custom Field Data - SPS Chapter-StudentLeadership-President Email'] = master_record['Incoming SPS President Email']
            # target_df.at[index, 'Custom Field Data - SPS Chapter-StudentLeadership-Vice President Name'] = master_record['Incoming SPS Vice President Name']
            # target_df.at[index, 'Custom Field Data - SPS Chapter-StudentLeadership-Vice President Email'] = master_record['Incoming SPS Vice President Email']
            # target_df.at[index, 'Custom Field Data - SPS Chapter-StudentLeadership-Secretary Name'] = master_record['Incoming SPS Secretary Name']
            # target_df.at[index, 'Custom Field Data - SPS Chapter-StudentLeadership-Secretary Email'] = master_record['Incoming SPS Secretary Email']
            # target_df.at[index, 'Custom Field Data - SPS Chapter-StudentLeadership-Treasurer Name'] = master_record['Incoming SPS Treasurer Name']
            # target_df.at[index, 'Custom Field Data - SPS Chapter-StudentLeadership-Treasurer Email'] = master_record['Incoming SPS Treasurer Email']
            # target_df.at[index, 'Custom Field Data - SPS Chapter-StudentLeadership-Other Officers Names'] = master_record['Other Officers (Format: Name_1; Title_1; Name_2; Title_2 )']
            # target_df.at[index, 'Custom Field Data - SPS Chapter-StudentLeadership-Other Officers Emails'] = master_record['Other Officers Email (Format: email1@mail.edu; email2@mail.edu)']

            # # Update the induction date separately
            # target_df = update_induction_date(school_name, city, target_df, index, induction_df)

    # Return the updated target dataframe
    return target_df

# Loop through each record in the master dataframe and update the target dataframe accordingly
for i, row in master_df.iterrows():
    target_df = update_record(row, target_df, induction_df)

# Save the updated target dataframe to a new Excel file
target_df.to_excel('Updated Zone 1 Activity New.xlsx', index=False)


# # Testing
school_name = "Harvard College"
school_names = ["Harvard University"]
# # school_name = "University of Massachusetts - Amherst"
# # school_names = ["University of Massachusetts Amherst"]
best_match, score = process.extractOne(school_name, school_names)
print(f"Best match: {best_match}, Score: {score}")
