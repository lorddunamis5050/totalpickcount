# hourly_pick_totals.py

import openpyxl
import pandas as pd
from openpyxl.utils.dataframe import dataframe_to_rows


START_TIME_PUTWALL = pd.to_datetime('8:00 PM', format='%I:%M %p')
END_TIME_PUTWALL = pd.to_datetime('11:59 PM', format='%I:%M %p')

def calculate_regular_pick_totals(df):
    # Initialize a dictionary to store regular pick totals per hour
    regular_pick_totals = {hour: 0 for hour in range(20, 24)}

    for index, row in df.iterrows():
        action = row['Action']
        datetime = row['DateTime']

        if action == 'REGULAR PICK':
            hour = datetime.hour

            if hour >= 20 and hour < 24:
                regular_pick_totals[hour] += row['Quantity']

    return regular_pick_totals

def calculate_single_pick_totals(df):
    # Initialize a dictionary to store single pick totals per hour
    single_pick_totals = {hour: 0 for hour in range(20, 24)}

    for index, row in df.iterrows():
        action = row['Action']
        packslip = row['Packslip']
        datetime = row['DateTime']

        if action == 'REPLNISH' and len(str(packslip)) >= 7 and str(packslip)[6] == 'S' and datetime >= START_TIME_PUTWALL and datetime <= END_TIME_PUTWALL:
            hour = datetime.hour

            if hour >= 20 and hour < 24:
                single_pick_totals[hour] += row['Quantity']

    return single_pick_totals

def calculate_replenishment_pick_totals(df):
    # Initialize a dictionary to store replenishment pick totals per hour
    replenishment_pick_totals = {hour: 0 for hour in range(20, 24)}

    for index, row in df.iterrows():
        action = row['Action']
        packslip = row['Packslip']
        datetime = row['DateTime']

        if action == 'REPLNISH' and len(str(packslip)) >= 7 and str(packslip)[6] == 'P' and datetime >= START_TIME_PUTWALL and datetime <= END_TIME_PUTWALL:
            hour = datetime.hour

            if hour >= 20 and hour < 24:
                replenishment_pick_totals[hour] += row['Quantity']

    return replenishment_pick_totals

def calculate_putwall_pick_totals(df):

    # Initialize a dictionary to store putwall pick totals per hour
    putwall_pick_totals = {hour: 0 for hour in range(20, 24)}

    for index, row in df.iterrows():
        action = row['Action']
        bin_label = row['BinLabel']

        if action == 'PUTWALL PICKING' and bin_label.startswith('MW'):
            hour = row['DateTime'].hour

            if hour >= 20 and hour < 24:
                putwall_pick_totals[hour] += row['Quantity']

    return putwall_pick_totals

def perform_hourly_pick_totals_analysis(df, book):
    # Calculate Regular Pick totals per hour
    regular_pick_totals = calculate_regular_pick_totals(df)

    # Calculate Single Pick totals per hour
    single_pick_totals = calculate_single_pick_totals(df)

    # Calculate Replenishment Pick totals per hour
    replenishment_pick_totals = calculate_replenishment_pick_totals(df)

    # Calculate Putwall Pick totals per hour
    putwall_pick_totals = calculate_putwall_pick_totals(df)

    # Create the "Total Units picked by hour" sheet if it doesn't exist
    if 'Total Units picked by hour' not in book.sheetnames:
        hourly_pick_totals_sheet = book.create_sheet('Total Units picked by hour')
    else:
        hourly_pick_totals_sheet = book['Total Units picked by hour']

    # Write the header row
    header_row = ['Hour', 'Regular Pick', 'Single Pick', 'Replenishment Pick', 'Putwall Pick']
    hourly_pick_totals_sheet.append(header_row)

    # Calculate the total quantity for each hour and write to the sheet
    for hour in range(20, 24):
        regular_pick_quantity = regular_pick_totals.get(hour, 0)
        single_pick_quantity = single_pick_totals.get(hour, 0)
        replenishment_pick_quantity = replenishment_pick_totals.get(hour, 0)
        putwall_pick_quantity = putwall_pick_totals.get(hour, 0)

        # Create a row with quantities for each pick type
        hour_data = [hour, regular_pick_quantity, single_pick_quantity, replenishment_pick_quantity, putwall_pick_quantity]
        hourly_pick_totals_sheet.append(hour_data)

    # Calculate the total quantities for each pick type
    total_regular_pick = sum(regular_pick_totals.values())
    total_single_pick = sum(single_pick_totals.values())
    total_replenishment_pick = sum(replenishment_pick_totals.values())
    total_putwall_pick = sum(putwall_pick_totals.values())

    # Add the totals row to the sheet
    total_row = ['Total', total_regular_pick, total_single_pick, total_replenishment_pick, total_putwall_pick]
    hourly_pick_totals_sheet.append(total_row)

    # Save the Excel file
    output_excel_file = 'pick_counts.xlsx'
    book.save(output_excel_file)

    print("Hourly pick totals analysis completed and saved.")
