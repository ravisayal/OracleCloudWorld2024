import json
import pandas as pd
import glob
import os

# Define the directory and the file pattern
directory = r'G:\My Drive\00. Work\04_Oracle_Delivery_lead\14_Oracle_Cloud_World_2024'
file_pattern = os.path.join(directory, 'output_*.json')

# List to store all the session details
all_sessions = []

# Get a list of all files matching the pattern
files = glob.glob(file_pattern)

# Parse each file
for file_path in files:
    try:
        # Specify the encoding, using 'utf-8' for Unicode files
        with open(file_path, 'r', encoding='utf-8') as file:
            data = json.load(file)

            # Check if 'items' exists in the JSON structure
            if 'items' in data and len(data['items']) > 0:
                for i in range(len(data['items'])):
                    session = data['items'][i]

                    # Extract participants' names and join them as comma-separated strings
                    participants = [participant['globalFullName'] for participant in session.get('participants', [])]
                    participants_str = ', '.join(participants) if participants else 'N/A'

                    # Collect session information, including externalID
                    session_info = {
                        "code": session['code'],
                        "type": session['type'],
                        "title": session['title'],
                        "abstract": session.get('abstract', 'N/A'),
                        "date": session['times'][0]['date'] if session.get('times') else 'N/A',
                        "startTimeFormatted": session['times'][0]['startTimeFormatted'] if session.get(
                            'times') else 'N/A',
                        "endTimeFormatted": session['times'][0]['endTimeFormatted'] if session.get('times') else 'N/A',
                        "externalID": session.get('externalID', 'N/A'),
                        "participants": participants_str
                    }

                    # Loop through attributes and add them as separate columns
                    attributes = session.get('attributevalues', [])
                    for index, attribute in enumerate(attributes):
                        attribute_column_name = attribute.get('attribute', f'attribute_{index + 1}_name')
                        session_info[attribute_column_name] = attribute.get('value', 'N/A')

                    # Append the session info to the list
                    all_sessions.append(session_info)
            else:
                print(f"'items' not found or empty in file: {file_path}")

    except UnicodeDecodeError as e:
        print(f"Error decoding file {file_path}: {e}")
    except json.JSONDecodeError as e:
        print(f"Error parsing JSON in file {file_path}: {e}")

# Convert the list to a DataFrame if there is data to save
if all_sessions:
    df_all_sessions = pd.DataFrame(all_sessions)

    # Save the combined data to an Excel file
    output_file = os.path.join(directory, 'combined_output.xlsx')
    df_all_sessions.to_excel(output_file, index=False)

    print(f"Data successfully parsed and saved to {output_file}")
else:
    print("No valid session data found to save.")
