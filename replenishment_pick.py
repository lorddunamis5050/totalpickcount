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



         # Find the highest time worked by any user within the time range
    highest_hours_worked = (replenishment_pick_df.groupby('UserID')['DateTime']
                             .agg(lambda x: (x.max() - x.min()).total_seconds() / 3600)
                             .max())

    # Count the sum of 'Quantity' for "REPLENISHMENT PICK" actions per user within the time range for PUTWALL PICKING
    replenishment_pick_per_user = replenishment_pick_df.groupby('UserID')['Quantity'].sum().reset_index(name='ReplenishmentPickQuantity')

            # Calculate UPH for each user, using the highest time as the denominator
    replenishment_pick_per_user['UPH'] = replenishment_pick_per_user['ReplenishmentPickQuantity'] / highest_hours_worked

            # Convert both "PutwallPickingQuantity" and "UPH" values to their absolute values
    replenishment_pick_per_user['ReplenishmentPickQuantity'] = abs(replenishment_pick_per_user['ReplenishmentPickQuantity'])
    replenishment_pick_per_user['UPH'] = abs(replenishment_pick_per_user['UPH'])


            # Create the "REPLENISHMENT  PICKING" sheet if it doesn't exist
    if 'REPLENISHMENT PICK' not in book.sheetnames:
        replenishment_pick_sheet = book.create_sheet('REPLENISHMENT PICK')
    else:
        replenishment_pick_sheet = book['REPLENISHMENT PICK']

                    # Write the header row
    header_row = ['UserID', 'ReplenishmentPickQuantity', 'UPH']
    replenishment_pick_sheet.append(header_row)

            # Convert the DataFrame to a list of lists for writing to Excel
    replenishment_picking_data = replenishment_pick_per_user[['UserID', 'ReplenishmentPickQuantity', 'UPH']].values.tolist()

        # Write the data to the Excel sheet
    for row_data in replenishment_picking_data:
        replenishment_pick_sheet.append(row_data)


    print("ByGroup PICKING analysis completed.")


