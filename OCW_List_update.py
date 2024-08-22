import pandas as pd
import re

# Define file paths
file_path = r'G:\My Drive\00. Work\04_Oracle_Delivery_lead\14_Oracle_Cloud_World_2024\Oracle_Cloud_World_2024_sessions.xlsx'
output_path = r'G:\My Drive\00. Work\04_Oracle_Delivery_lead\14_Oracle_Cloud_World_2024\Oracle_Cloud_World_2024_sessions_modified.xlsx'

# Load all sheets into a dictionary of DataFrames without treating the first row as header
sheets = pd.read_excel(file_path, sheet_name=None, header=None)

# Define the header row
header_row = ['Title', 'Summary', 'Presenter', 'Schedule']

# Step 1: Clean up rows with the unwanted text in the first cell and shift cells to the left
for sheet_name, df in sheets.items():
    for index, row in df.iterrows():
        if str(row[0]).strip() == "Recommended For You Rate this recommendation thumbs-up thumbs-down":
            # Delete the first cell and shift cells in the row to the left
            df.iloc[index, :] = df.iloc[index, :].shift(-1)
            print(f"Row {index + 1} in sheet {sheet_name} has been cleaned.")


# Define a function to check if a value is a valid custom date format with optional leading space
def is_custom_date_format(value):
    if pd.isnull(value):
        return False

    # Pattern with optional space at the beginning
    pattern = r"^\s*(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday), [A-Za-z]{3} \d{1,2}"

    # Match the pattern to the value
    if re.match(pattern, str(value)):
        return True
    return False


# Function to split the schedule column into Day, Date, Time Slot PDT, Time Slot EDT
def split_schedule(schedule):
    pattern = r"^(?P<Day>\w+), (?P<Date>[A-Za-z]{3} \d{1,2}) (?P<Time_Slot_PDT>[\d:APM\s-]+) PDT \( (?P<Time_Slot_EDT>[\d:APM\s-]+) EDT \)$"
    match = re.match(pattern, schedule)
    if match:
        return match.group("Day"), match.group("Date"), match.group("Time_Slot_PDT"), match.group("Time_Slot_EDT")
    return None, None, None, None


# Function to extract session ID from Column A
def extract_session_id(title):
    pattern = r"\[(.*?)\]$"
    match = re.search(pattern, title)
    if match:
        return match.group(1)
    return None


# Function to convert the "Time Slot PDT" into a proper datetime format for sorting
def convert_time_slot_to_datetime(time_slot):
    try:
        start_time = re.search(r"(\d{1,2}:\d{2} (AM|PM))", time_slot).group(1)
        return pd.to_datetime(start_time, format="%I:%M %p")
    except:
        return None


# Process each sheet
for sheet_name, df in sheets.items():
    print(f"Processing sheet: {sheet_name}")

    # Add the header row to the dataframe
    df.columns = header_row + list(df.columns[len(header_row):])  # Add the new header to the beginning

    # Loop through rows in the sheet
    for index, row in df.iterrows():
        print(f"Processing row {index + 1}")

        # Extract session ID from Column 'Title'
        session_id = extract_session_id(row['Title'])
        df.at[index, 'Session ID'] = session_id

        # Trim leading spaces from Column 'Schedule'
        df.at[index, 'Schedule'] = str(row['Schedule']).strip()

        # Split the schedule column into Day, Date, Time Slot PDT, Time Slot EDT
        schedule_value = df.at[index, 'Schedule']
        day, date, time_pdt, time_edt = split_schedule(schedule_value)

        # Add the split columns to the dataframe
        df.at[index, 'Day'] = day
        df.at[index, 'Date'] = date
        df.at[index, 'Time Slot PDT'] = time_pdt
        df.at[index, 'Time Slot EDT'] = time_edt

    # Step 2: Delete any column where all values are empty
    df.dropna(axis=1, how='all', inplace=True)

    # Step 3: Convert 'Time Slot PDT' to datetime for sorting
    df['Time Slot PDT (datetime)'] = df['Time Slot PDT'].apply(convert_time_slot_to_datetime)

    # Step 4: Sort the dataframe by 'Date', 'Time Slot PDT (datetime)', and 'Session ID'
    df.sort_values(by=['Date', 'Time Slot PDT (datetime)', 'Session ID'], ascending=[True, True, True], inplace=True)

    # Step 5: Drop the temporary 'Time Slot PDT (datetime)' column after sorting
    df.drop(columns=['Time Slot PDT (datetime)'], inplace=True)

# Save the updated Excel file
print("Saving the modified file...")
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    for sheet_name, df in sheets.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False, header=True)  # Ensure headers are included

print("Processing complete.")
