import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows

# Define the time range
START_TIME = pd.to_datetime('8:00 PM', format='%I:%M %p')
END_TIME = pd.to_datetime('11:59 PM', format='%I:%M %p')

def determine_zone(row):
    if row['action'] == 'PICKLINE':
        if row['bin_label'].startswith(('1H', '1G', '2E', '2H', '3F', '3H', '3R', '2R', '1Y', '1C', '1D', '2D', '3D')) and not row['packslip'].startswith('TR'):
            return row['bin_label'][:2]  # Assuming the zone is the first two characters of bin_label
    return None  # Return None or a default value if the row doesn't match the criteria

# Make sure 'action', 'bin_label', and 'packslip' are columns in your DataFrame
# df['Zone'] = df.apply(determine_zone, axis=1)


def calculate_picks_per_zone(df, zones):
    # Filter the DataFrame for the given zones and sum quantities
    pick_totals_per_zone = df[df['Zone'].isin(zones)].groupby('Zone')['Quantity'].sum().to_dict()
    return pick_totals_per_zone

def perform_pick_totals_analysis_per_zones(df, book):

    df['Zone'] = df.apply(determine_zone, axis=1)
    # Define the pick types and corresponding action filters
    pick_types = {
        'Regular Pick': ['REGULAR PICK'],
        'Single Pick': ['SINGLE PICK'],
        'Replenishment Pick': ['REPLENISHMENT PICK'],
        'Putwall Pick': ['PUTWALL PICKING']
    }

    zones = ['1H', '1G', '2E', '2H', '3F', '3H', '3R', '2R', '1Y', '1C', '1D', '2D', '3D']

    # Calculate total picks per zone
    pick_totals_per_zone = calculate_picks_per_zone(df, zones)

    # Adding to a workbook
    zone_df = pd.DataFrame.from_dict(pick_totals_per_zone, orient='index', columns=['Total Picks'])
    zone_sheet = book.create_sheet(title="Zone Analysis")
    for row in dataframe_to_rows(zone_df, index=True, header=True):
        zone_sheet.append(row)


