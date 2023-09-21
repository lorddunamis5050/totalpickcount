import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows

def perform_replenishment_pick_analysis(df, book):
    # Define your desired time range for PUTWALL PICKING
    START_TIME_PUTWALL = pd.to_datetime('8:00 PM', format='%I:%M %p')
    END_TIME_PUTWALL = pd.to_datetime('11:59 PM', format='%I:%M %p')

    # Initialize DataFrames to store replenishment pick data
    replenishment_pick_per_user = pd.DataFrame(columns=['UserID', 'ReplenishmentPickQuantity'])

    # Function to check for "REPLENISHMENT PICK"
    def is_replenishment_pick(row):
        action = row['Action']
        packslip = row['Packslip']
        datetime = row['DateTime']

        if action == 'REPLNISH' and len(str(packslip)) >= 7 and str(packslip)[6] == 'P' and datetime >= START_TIME_PUTWALL and datetime <= END_TIME_PUTWALL:
            return True

        return False

    # Apply the function to the DataFrame to identify "REPLENISHMENT PICK"
    df['IsReplenishmentPick'] = df.apply(is_replenishment_pick, axis=1)

    # Filter rows based on the criteria for "REPLENISHMENT PICK"
    replenishment_pick_df = df[df['IsReplenishmentPick']]

    # Count the sum of 'Quantity' for "REPLENISHMENT PICK" actions per user within the time range for PUTWALL PICKING
    replenishment_pick_per_user = replenishment_pick_df.groupby('UserID')['Quantity'].sum().reset_index(name='ReplenishmentPickQuantity')

    # Create a new Excel workbook (ensure you have the 'book' variable defined in the main script)
    replenishment_pick_sheet = book.create_sheet('REPLENISHMENT PICK')

    # Write the REPLENISHMENT PICK data to an Excel sheet
    for row_data in dataframe_to_rows(replenishment_pick_per_user, index=False, header=True):
        replenishment_pick_sheet.append(row_data)
