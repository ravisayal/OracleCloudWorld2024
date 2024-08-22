import json
import pandas as pd
import glob
import os

# Define the directory and the file pattern
directory = r'G:\My Drive\00. Work\04_Oracle_Delivery_lead\14_Oracle_Cloud_World_2024'
file_pattern = os.path.join(directory, 'output_0.json')

# List to store all the session details
all_sessions = []

# Get a list of all files matching the pattern
files = glob.glob(file_pattern)

# Column sequence to maintain
column_sequence = [
    "code", "type", "title", "abstract", "date", "startTimeFormatted", "endTimeFormatted",
    "externalID", "participants", "Content Area", "Area of Interest", "Industry",
    "Session Format", "Audience Level", "Job Role", "Session Type", "Mobile Published",
    "Day", "Time", "OCW Recommendations Topics", "NOTE", "Scheduling", "Session Settings"
]

# Parse each file
for file_path in files:
    try:
        # Specify the encoding, using 'utf-8' for Unicode files
        with open(file_path, 'r', encoding='utf-8') as file:
            data = json.load(file)

            # Check if 'sectionList' exists in the JSON structure
            if 'sectionList' in data and len(data['sectionList']) > 0:
                for section in data['sectionList']:
                    for i in range(len(section['items'])):
                        session = section['items'][i]

                        # Extract participants' names and join them as comma-separated strings
                        participants = [participant['globalFullName'] for participant in
                                        session.get('participants', [])]
                        participants_str = ', '.join(participants) if participants else 'N/A'

                        # Collect session information in the given column sequence, with placeholders for missing fields
                        session_info = {
                            "code": session.get('code', 'N/A'),
                            "type": session.get('type', 'N/A'),
                            "title": session.get('title', 'N/A'),
                            "abstract": session.get('abstract', 'N/A'),
                            "date": session['times'][0]['date'] if session.get('times') else 'N/A',
                            "startTimeFormatted": session['times'][0]['startTimeFormatted'] if session.get(
                                'times') else 'N/A',
                            "endTimeFormatted": session['times'][0]['endTimeFormatted'] if session.get(
                                'times') else 'N/A',
                            "externalID": session.get('externalID', 'N/A'),
                            "participants": participants_str,
                            "Content Area": "N/A",  # Default placeholder
                            "Area of Interest": "N/A",  # Default placeholder
                            "Industry": "N/A",  # Default placeholder
                            "Session Format": "N/A",  # Default placeholder
                            "Audience Level": "N/A",  # Default placeholder
                            "Job Role": "N/A",  # Default placeholder
                            "Session Type": "N/A",  # Default placeholder
                            "Mobile Published": "N/A",  # Default placeholder
                            "Day": "N/A",  # Default placeholder
                            "Time": "N/A",  # Default placeholder
                            "OCW Recommendations Topics": "N/A",  # Default placeholder
                            "NOTE": "N/A",  # Default placeholder
                            "Scheduling": "N/A",  # Default placeholder
                            "Session Settings": "N/A"  # Default placeholder
                        }

                        # Loop through attributes and add them in the specified fields
                        attributes = session.get('attributevalues', [])
                        for attribute in attributes:
                            attribute_name = attribute.get('attribute', 'N/A')
                            attribute_value = attribute.get('value', 'N/A')

                            # Assign values based on matching attribute names
                            if attribute_name in session_info:
                                session_info[attribute_name] = attribute_value

                        # Append the session info to the list
                        all_sessions.append(session_info)
            else:
                print(f"'sectionList' not found or empty in file: {file_path}")

    except UnicodeDecodeError as e:
        print(f"Error decoding file {file_path}: {e}")
    except json.JSONDecodeError as e:
        print(f"Error parsing JSON in file {file_path}: {e}")

# Convert the list to a DataFrame if there is data to save
if all_sessions:
    df_all_sessions = pd.DataFrame(all_sessions)

    # Ensure the columns are in the specified order
    df_all_sessions = df_all_sessions[column_sequence]

    # Save the combined data to an Excel file
    output_file = os.path.join(directory, 'combined_output_pt1.xlsx')
    df_all_sessions.to_excel(output_file, index=False)

    print(f"Data successfully parsed and saved to {output_file}")
else:
    print("No valid session data found to save.")
