# SPS Report Automation

## Abstract

Effective management and support of physics chapters across the US is crucial for fostering engagement and development within the Society of Physics Students (SPS). As the SPS Program Engagement Summer Intern at the American Institute of Physics, I play a key role in overseeing the collection and review of chapter reports from approximately 200 school chapters. These reports provide vital information on chapter activities, leadership updates, and engagement levels. My responsibilities include reviewing each submission, ensuring the accuracy and currency of records in zone-specific databases, and preparing comprehensive activity reports. These reports are essential for the SPS-AlP main office to understand chapter status and strategize support accordingly.

To enhance the efficiency of this process, I am exploring automation options, including Python scripting, to streamline data management. This exploration aims to determine whether automation can improve the accuracy and efficiency of handling chapter reports or if manual processing remains the optimal approach. My efforts contribute significantly to the national office's capacity to support and engage SPS chapters effectively.

## Overview

The **SPS Report Automation** project aims to improve the efficiency of managing and reviewing chapter reports for the Society of Physics Students (SPS). The project utilizes Python to automate data handling tasks, including fuzzy matching of school names, updating records with new information, and synchronizing data across multiple Excel files.

## Features

- **Data Integration**: Reads and integrates data from multiple Excel files.
- **Fuzzy Matching**: Uses fuzzy matching to handle discrepancies in school names and ensure accurate data updates.
- **Data Update**: Updates target dataframes with new information, including chapter reports and induction dates.
- **Flexibility**: Allows for manual overrides and customizable matching thresholds.

## Script Functions

- **get_correct_match**: Finds the best match for a school name using fuzzy matching and manual overrides.
- **update_record**: Updates the target dataframe with details from the master record based on school name matching.
- **update_chapter_reports**: Updates chapter reports with current year data, handling duplicate entries and ensuring consistency.
- **update_induction_date**: Updates the induction date for schools based on fuzzy matching with an induction date file.

## Setup Instructions

1. **Clone the Repository**

   ```sh
   git clone https://github.com/Collins-kariuk/sps-report-automation.git
   cd sps-report-automation
   ```
   
2. **Install Dependencies**

   Ensure you have the required Python libraries installed. You can install them using pip:

   ```sh
   pip install pandas fuzzywuzzy openpyxl
   ```

3. **Prepare Your Data**

   Ensure you have the following Excel files in the project directory:
   - `24 Reports Zone 1.xlsx` (Master file with chapter reports)
   - `Zone 1 Activity.xlsx` (Target file to be updated)
   - `MHS Chapters.xlsx` (Induction dates file)

4. **Run the Script**

   Execute the script to perform data updates:

   ```sh
   python your_script_name.py
   ```

   Make sure to replace `your_script_name.py` with the actual name of your Python script.

5. **Review Output**

   The script will save the updated target dataframe to a new Excel file. Review the output file to ensure the updates are accurate.

## Technologies Used

- **Python**: For scripting and automation.
- **Pandas**: For data manipulation and analysis.
- **FuzzyWuzzy**: For fuzzy matching of school names.

## Contributing

1. **Fork the Repository**

   Create a personal fork of the repository to make changes.

2. **Create a Branch**

   Create a new branch for your changes.

   ```sh
   git checkout -b feature/new-feature
   ```

3. **Make Changes**

   Implement your feature or bug fix.

4. **Commit Changes**

   Commit your changes with a descriptive message.

   ```sh
   git commit -am 'Add new feature'
   ```

5. **Push to GitHub**

   Push your changes to your fork.

   ```sh
   git push origin feature/new-feature
   ```

6. **Create a Pull Request**

   Open a pull request from your branch to the main repository.

## License

This project is licensed under the MIT License.

## Contact

For any questions or suggestions, please reach out to me via my profile.
