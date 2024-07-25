import unittest
import pandas as pd
from Automate import update_record

class TestAutomation(unittest.TestCase):

    def setUp(self):
        # Sample data to mimic the actual dataframes
        self.master_df = pd.DataFrame({
            'School Name (No abbreviations please)': [
                'Harvard College ', 'University of Massachusetts Amherst', 'Fairfield University'
            ],
            'Chapter Adviser Name': [
                'Dr. David Morin', 'Ben Heidenreich', 'Angela Biselli and Robert Cordery '
            ],
            'Chapter Adviser Email': [
                'djmorin@fas.harvard.edu', 'bheidenreich@umass.edu', 'rcordery@fairfield.edu'
            ]
        })

        self.target_df = pd.DataFrame({
            'Custom Field Data - Chapter School Name': [
                'Harvard University', 'University of Massachusetts-Amherst', 'Fairfield University'
            ],
            'Custom Field Data - SPS Chapter-Advisor Name': [
                '', '', ''
            ],
            'Custom Field Data - SPS Chapter-Advisor Email': [
                '', '', ''
            ]
        })

    def test_update_record(self):
        updated_target_df = self.target_df.copy()

        for _, master_record in self.master_df.iterrows():
            updated_target_df = update_record(master_record, updated_target_df)

        # Check if Harvard University's chapter advisor details were correctly updated
        # self.assertEqual(
        #     updated_target_df.loc[
        #         updated_target_df['Custom Field Data - Chapter School Name'] == 'Harvard University',
        #         'Custom Field Data - SPS Chapter-Advisor Name'
        #     ].values[0],
        #     'Dr. David Morin'
        # )
        # self.assertEqual(
        #     updated_target_df.loc[
        #         updated_target_df['Custom Field Data - Chapter School Name'] == 'Harvard University',
        #         'Custom Field Data - SPS Chapter-Advisor Email'
        #     ].values[0],
        #     'djmorin@fas.harvard.edu'
        # )

        # Check if University of Massachusetts Amherst's chapter advisor details were correctly updated
        self.assertEqual(
            updated_target_df.loc[
                updated_target_df['Custom Field Data - Chapter School Name'] == 'University of Massachusetts-Amherst',
                'Custom Field Data - SPS Chapter-Advisor Name'
            ].values[0],
            'Ben Heidenreich'
        )
        self.assertEqual(
            updated_target_df.loc[
                updated_target_df['Custom Field Data - Chapter School Name'] == 'University of Massachusetts-Amherst',
                'Custom Field Data - SPS Chapter-Advisor Email'
            ].values[0],
            'bheidenreich@umass.edu'
        )

        # Check if Fairfield University's chapter advisor details were correctly updated
        self.assertEqual(
            updated_target_df.loc[
                updated_target_df['Custom Field Data - Chapter School Name'] == 'Fairfield University',
                'Custom Field Data - SPS Chapter-Advisor Name'
            ].values[0],
            'Angela Biselli and Robert Cordery '
        )
        self.assertEqual(
            updated_target_df.loc[
                updated_target_df['Custom Field Data - Chapter School Name'] == 'Fairfield University',
                'Custom Field Data - SPS Chapter-Advisor Email'
            ].values[0],
            'rcordery@fairfield.edu'
        )

if __name__ == '__main__':
    unittest.main()
