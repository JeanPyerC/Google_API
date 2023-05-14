import pandas as pd
import openpyxl
import googlemaps
from datetime import datetime
from urllib3.exceptions import ProtocolError
import time

# Read the "Report" sheet of the "Routes_Request.xlsx" Excel file
df = pd.read_excel('Routes_Request.xlsx', sheet_name='Report')

# Extract the Origin-Full Address and Destination-Full Address columns as lists
origins = df['Origin-Full Address'].tolist()
destinations = df['Destination-Full Address'].tolist()

# Reading the KPI file
with open('Google_Maps_API-KEY.txt', 'r') as f:
    api_key = f.readline()

#Using Google Maps to access it
gmaps = googlemaps.Client(api_key)

# Creating a variable to hold distance & durations
distances = []
durations = []

# Creating a rule for a number of Reteries and delays
MAX_RETRIES = 3
DELAY_SECONDS = 2

# Creating a loop
for origin, destination in zip(origins, destinations):
    retries = 0
    while retries < MAX_RETRIES:
        try:
            result = gmaps.distance_matrix(origin, destination, mode="driving", units="imperial")
            distance = result["rows"][0]["elements"][0]["distance"]["text"]
            duration = result["rows"][0]["elements"][0]["duration"]["text"]
            print(f"Origin: {origin}, Destination: {destination}, Distance: {distance}, Duration: {duration}")

            distances.append(distance)
            durations.append(duration)
            break
        except (KeyError, ProtocolError) as e:
            if retries == MAX_RETRIES - 1:
                print(f"Max retries exceeded for request: {origin} to {destination}")
                distances.append("Not found")
                durations.append("Not found")
                break
            else:
                print(f"Error: {e}")
                retries += 1
                print(f"Retrying after {DELAY_SECONDS}s...")
                time.sleep(DELAY_SECONDS)
                continue


# Add the new columns to the DataFrame
df['Distance'] = distances
df['Duration'] = durations

# Save the updated DataFrame to a new Excel file with current date and time added to the filename
now = datetime.now()
dt_string = now.strftime("%Y%m%d_%H%M%S")
filename = f"updated_routes_{dt_string}.xlsx"
df.to_excel(filename, sheet_name='Report', index=False)
