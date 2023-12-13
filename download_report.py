import requests
import base64

# Set up authentication headers
username = "your_username"
password = "your_password"
auth_header = "Basic " + base64.b64encode(f"{username}:{password}".encode()).decode()
headers = {"Authorization": auth_header}

# Set up API parameters
params = {
    "action": "list",
    "status": "Finished",  # Only retrieve finished reports
    "type": "Scan"  # Only retrieve scan reports (not compliance, etc.)
}

# Send the API request and parse the response
url = "https://qualysapi.qg3.apps.qualys.eu/api/2.0/fo/report/"
response = requests.get(url, headers=headers, params=params)
report_ids = [line.split(",")[2] for line in response.text.split("\n")[2:-2]]
print(report_ids)


import requests
import base64

username = "your_username"
password = "your_password"
report_id = "your_report_id"

# Authenticate to the API
auth_header = "Basic " + base64.b64encode(f"{username}:{password}".encode()).decode()
headers = {"Authorization": auth_header}

# Get a list of available reports
list_url = "https://qualysapi.qg3.apps.qualys.eu/api/2.0/fo/report/"
list_params = {"action": "list", "status": "Finished"}
list_response = requests.get(list_url, headers=headers, params=list_params)

# Find the report ID for the report you want to download
reports = list_response.text.split("\n")[2:-2]
report_map = {}
for report in reports:
    report_fields = report.split(",")
    report_map[report_fields[2]] = report_fields[0]

# Download the report
download_url = f"https://qualysapi.qg3.apps.qualys.eu/api/2.0/fo/report/" \
    f"download/?id={report_map[report_id]}&output_format=pdf"
download_response = requests.get(download_url, headers=headers)

# Save the report to a file
with open("report.pdf", "wb") as f:
    f.write(download_response.content)
