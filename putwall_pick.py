import openpyxl
import pandas as pd
from openpyxl.utils.dataframe import dataframe_to_rows
import numpy as np

def perform_putwall_pick_analysis(df, book):
    # Define your desired time range for PUTWALL PICKING
    start_time_putwall = pd.to_datetime('8:00 PM', format='%I:%M %p')
    end_time_putwall = pd.to_datetime('11:59 PM', format='%I:%M %p')


    df['DateTime'] = pd.to_datetime(df['DateTime'])

    # Initialize a DataFrame to store PUTWALL PICKING data per user
    putwall_picking_per_user = pd.DataFrame(columns=['UserID', 'PutwallPickingQuantity'])

    # Function to modify the 'Action' column based on 'BinLabel' for PUTWALL PICKING
    def modify_action_putwall(row):
        action = row['Action']
        bin_label = row['BinLabel']
        
        if action == 'PICKLINE' and bin_label.startswith('MW'):
            return 'PUTWALL PICKING'
        
        return action

    # Apply the function to the DataFrame for PUTWALL PICKING
    df['Action'] = df.apply(modify_action_putwall, axis=1)

  # Filter rows based on the specified time range for PUTWALL PICKING
    filtered_df_putwall = df[(df['DateTime'] >= start_time_putwall) & (df['DateTime'] <= end_time_putwall)]


    # Group by "UserID" and calculate total Units picked
    putwall_picking_per_user = filtered_df_putwall[filtered_df_putwall['Action'] == 'PUTWALL PICKING'].groupby('UserID').agg(
        PutwallPickingQuantity=('Quantity', 'sum')
    ).reset_index()

        # Function to calculate the duration of putwall picking actions
    def calculate_putwall_picking_time(group):
        group = group.sort_values('DateTime')
     # Duration in minutes for each row
        group['Duration'] = group['DateTime'].diff().shift(-1).dt.total_seconds().div(60)
        # Mark the rows that are not part of continuous putwall picking
        group['IsGap'] = ~group['Action'].eq('PUTWALL PICKING') | group['Action'].shift().ne('PUTWALL PICKING')
        # Cumulatively sum the gap marks to create unique session IDs
        group['SessionId'] = group['IsGap'].cumsum()
        # Filter out non-putwall picking rows and rows that start a new session
        putwall_sessions = group[group['Action'].eq('PUTWALL PICKING') & ~group['IsGap']]
        # Sum the durations for each session
        session_durations = putwall_sessions.groupby('SessionId')['Duration'].sum()
        return session_durations.sum()  # Return the total time across all sessions

    # Filter for putwall picking actions only
    putwall_picking_actions = filtered_df_putwall[filtered_df_putwall['Action'] == 'PUTWALL PICKING']

    # Calculate total putwall picking time for each user considering gaps
    total_putwall_picking_time = putwall_picking_actions.groupby('UserID').apply(calculate_putwall_picking_time).reset_index(name='TotalPutwallPickingMinutes')

    # Merge this time with the putwall_picking_per_user DataFrame
    putwall_picking_per_user = putwall_picking_per_user.merge(total_putwall_picking_time, on='UserID', how='left')


# Calculate UPH (Units Per Hour) for each user using the total putwall picking time
    putwall_picking_per_user['UPH'] = putwall_picking_per_user.apply(
    lambda row: (row['PutwallPickingQuantity'] * 60) / row['TotalPutwallPickingMinutes'] if row['TotalPutwallPickingMinutes'] >= 30 else 0, axis=1
    )

    # Calculate UPH for each user, using the highest time as the denominator
    # putwall_picking_per_user['UPH'] = putwall_picking_per_user['PutwallPickingQuantity'] / highest_hours_worked

        # Convert both "PutwallPickingQuantity" and "UPH" values to their absolute values
    putwall_picking_per_user['PutwallPickingQuantity'] = abs(putwall_picking_per_user['PutwallPickingQuantity'])
    putwall_picking_per_user['UPH'] = abs(putwall_picking_per_user['UPH'])

    # Calculate the average UPH, excluding zeros
    average_uph = putwall_picking_per_user.loc[putwall_picking_per_user['UPH'] > 0, 'UPH'].mean()

        # Replace NaN or infinite values with zero
    putwall_picking_per_user.replace([np.inf, -np.inf, np.nan], 0, inplace=True)


    # Create the "PUTWALL PICKING" sheet if it doesn't exist
    if 'PUTWALL PICKING' not in book.sheetnames:
        putwall_picking_sheet = book.create_sheet('PUTWALL PICKING')
    else:
        putwall_picking_sheet = book['PUTWALL PICKING']

    # Update the header row to include 'TotalPutwallPickingMinutes'
    header_row = ['UserID', 'PutwallPickingQuantity', 'TotalPutwallPickingMinutes', 'UPH']
    putwall_picking_sheet.append(header_row)

    # Convert the DataFrame to a list of lists for writing to Excel
    putwall_picking_data = putwall_picking_per_user[['UserID', 'PutwallPickingQuantity', 'TotalPutwallPickingMinutes', 'UPH']].values.tolist()

    # Write the data to the Excel sheet
    for row_data in putwall_picking_data:
        putwall_picking_sheet.append(row_data)
        
    # Write the average UPH row to the Excel sheet
    average_uph_row = ["Average UPH", "", "", average_uph]
    putwall_picking_sheet.append(average_uph_row)

    print("PUTWALL PICKING analysis completed.")

