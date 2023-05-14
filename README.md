# Google Maps API - Distance and Travel

One of my clients approached me with a request to gather distances between stops on their transportation routes. To accomplish this quickly and accurately, we decided to use the Google Maps API. While obtaining an API key is beyond the scope of this report, we'll explain how to use the API to collect the necessary data. We'd like to emphasize that sharing your API key can lead to unexpected costs, so it's important to keep it secure.

Since the client exports their transportation details to an Excel file, I created a script (application) to read the file, extract the required information, and create a new file containing the requested data. To make things easier for the client, I used PyInstaller to create a single executable file. This means that the client doesn't have to install any additional software, they can simply download the app, double-click it, and everything will run smoothly.

Here's a breakdown of what the code does:

### Step 1 - Import the necessary libraries: 

The script starts by importing the following libraries:
- pandas: a library for working with tabular data in Python, used here to read and write Excel files
- openpyxl: a library for reading and writing Excel files in Python
- googlemaps: a library for interacting with the Google Maps API
- datetime: a module for working with dates and times in Python
- urllib3.exceptions.ProtocolError: a module to handle Protocol errors

```
import pandas as pd
import openpyxl
import googlemaps
from datetime import datetime
from urllib3.exceptions import ProtocolError
import time
```

### Step 2 - Read the input data:

The script reads the "Report" sheet of an Excel file named "Routes_Request.xlsx" using pandas, and extracts the contents of the "Origin-Full Address" and "Destination-Full Address" columns as lists. These lists are used to store the starting and ending addresses for each route.

```
# Read the "Report" sheet of the "Routes_Request.xlsx" Excel file
df = pd.read_excel('Routes_Request.xlsx', sheet_name='Report')

# Extract the Origin-Full Address and Destination-Full Address columns as lists
origins = df['Origin-Full Address'].tolist()
destinations = df['Destination-Full Address'].tolist()
```

### Step 3 - Get the Googe API key:

The script reads the Google Maps API key from a text file named "Google_Maps_API-KEY.txt" and stores it in a variable called "api_key".

```
# Reading the KPI file
with open('Google_Maps_API-KEY.txt', 'r') as f:
    api_key = f.readline()
```

### Step 4 - Create a Google Maps client object:

The script creates a Google Maps client object using the api_key variable.

```
#Using Google Maps to access it
gmaps = googlemaps.Client(api_key)
```

### Step 5 - Loop over the pairs of addresses and handle exceptions:

The script creates a loop that iterates over pairs of starting and ending addresses. For each pair, the script makes a request to the Google Maps API to get the distance and duration between the two addresses using the "distance_matrix" function provided by the client object.

The script handles exceptions that may occur during the API request, such as a missing or invalid API key, a timeout error, or a protocol error. If an exception occurs, the script retries the request up to a maximum number of times (defined as MAX_RETRIES). If the maximum number of retries is reached without success, the script records "Not found" values for the distance and duration and moves on to the next pair of addresses.

```
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
```


### Step 6 - Save the results:

The script saves the distance and duration results for each pair of addresses back to the original Excel file by adding two new columns to the DataFrame called "Distance" and "Duration". It then saves the updated DataFrame to a new Excel file with a filename that includes the current date and time.

```
# Save the updated DataFrame to a new Excel file with current date and time added to the filename
now = datetime.now()
dt_string = now.strftime("%Y%m%d_%H%M%S")
filename = f"updated_routes_{dt_string}.xlsx"
df.to_excel(filename, sheet_name='Report', index=False)
```
