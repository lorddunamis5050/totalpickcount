import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows

def perform_regular_pick_analysis(df , book):
    # Define your desired time range for REGULAR PICK
    START_TIME_REGULAR = pd.to_datetime('8:00 PM', format='%I:%M %p')
    END_TIME_REGULAR = pd.to_datetime('11:59 PM', format='%I:%M %p')

    # Initialize DataFrames to store regular pick data
    regular_pick_per_user = pd.DataFrame(columns=['UserID', 'RegularPickQuantity'])

    # Function to modify the 'Action' column based on 'BinLabel' for REGULAR PICK
    def modify_action_regular(row):
        action = row['Action']
        bin_label = row['BinLabel']
        packslip = row['Packslip']

        if action == 'PICKLINE':
            if bin_label.startswith(('1H', '1G', '2E', '2H', '3F', '3H', '3R', '2R', '1Y', '1C', '1D', '2D', '3D')) and not packslip.startswith('TR'):
                return 'REGULAR PICK'

        return action

    # Apply the function to the DataFrame for REGULAR PICK
    df['Action'] = df.apply(modify_action_regular, axis=1)

    # Filter rows based on the specified time range for REGULAR PICK
    filtered_df_regular = df[(df['DateTime'] >= START_TIME_REGULAR) & (df['DateTime'] <= END_TIME_REGULAR)]

    # Count the sum of 'Quantity' for 'REGULAR PICK' actions per user within the time range for REGULAR PICK
    regular_pick_per_user = filtered_df_regular[filtered_df_regular['Action'] == 'REGULAR PICK'].groupby('UserID')['Quantity'].sum().reset_index(name='RegularPickQuantity')


    regular_pick_per_user['RegularPickQuantity'] = abs(regular_pick_per_user['RegularPickQuantity'])

    # Create a new Excel workbook (ensure you have the 'book' variable defined in the main script)
    regular_pick_sheet = book.create_sheet('REGULAR PICK')

    # Write the REGULAR PICK data to an Excel sheet
    for row_data in dataframe_to_rows(regular_pick_per_user, index=False, header=True):
        regular_pick_sheet.append(row_data)

    print("REGULAR PICKING analysis completed.")
