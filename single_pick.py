import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows

def perform_single_pick_analysis(df , book):
    # Define your desired time range for PUTWALL PICKING
    START_TIME_PUTWALL = pd.to_datetime('8:00 PM', format='%I:%M %p')
    END_TIME_PUTWALL = pd.to_datetime('11:59 PM', format='%I:%M %p')

    # Initialize DataFrames to store single pick data
    single_pick_per_user = pd.DataFrame(columns=['UserID', 'SinglePickQuantity'])

    # Function to check for "SINGLE PICK"
    def is_single_pick(row):
        action = row['Action']
        packslip = row['Packslip']
        datetime = row['DateTime']

        if action == 'REPLNISH' and len(str(packslip)) >= 7 and str(packslip)[6] == 'S' and datetime >= START_TIME_PUTWALL and datetime <= END_TIME_PUTWALL:
            return True

        return False

    # Apply the function to the DataFrame to identify "SINGLE PICK"
    df['IsSinglePick'] = df.apply(is_single_pick, axis=1)

    # Filter rows based on the criteria for "SINGLE PICK"
    single_pick_df = df[df['IsSinglePick']]

    # Count the sum of 'Quantity' for "SINGLE PICK" actions per user within the time range for PUTWALL PICKING
    single_pick_per_user = single_pick_df.groupby('UserID')['Quantity'].sum().reset_index(name='SinglePickQuantity')

    # Create a new Excel workbook (ensure you have the 'book' variable defined in the main script)
    single_pick_sheet = book.create_sheet('SINGLE PICK')

    # Write the SINGLE PICK data to an Excel sheet
    for row_data in dataframe_to_rows(single_pick_per_user, index=False, header=True):
        single_pick_sheet.append(row_data)
