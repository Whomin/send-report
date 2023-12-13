import requests
import base64
import xml.etree.ElementTree as ET
import json
import csv
import os

# Replace <QUALYS_API_SERVER>, <QUALYS_API_USERNAME>, and <QUALYS_API_PASSWORD>
# with your Qualys API server, username, and password, respectively.
api_server = "https://qualysapi.qualys.com"
username = "tfs_st2"
password = "Toyota11!"
#password = os.environ.get('QUALYS_PASS', 'Default Value')

# Build the API request URL
url = f"{api_server}/api/2.0/fo/report/?action=list "
download_url = f"{api_server}/api/2.0/fo/report/?action=fetch&id="

# Build the API request headers
headers = {
    "X-Requested-With": "Python",
    "Content-Type": "application/xml",
    "Authorization": f"Basic {base64.b64encode(bytes(f'{username}:{password}', 'utf-8')).decode('utf-8')}"
}

# Send the API request and retrieve the response
response = requests.post(url, headers=headers)

xml_data = response.text

def xml_to_dict(element):
    element_dict = element.attrib
    for child in element:
        child_dict = xml_to_dict(child)
        if child.tag in element_dict:
            if isinstance(element_dict[child.tag], list):
                element_dict[child.tag].append(child_dict)
            else:
                element_dict[child.tag] = [element_dict[child.tag], child_dict]
        else:
            element_dict[child.tag] = child_dict
    if not element:
        element_dict = element.text or ''
    return element_dict

parsed_json = xml_to_dict(ET.fromstring(xml_data))
reports = parsed_json['RESPONSE']['REPORT_LIST']['REPORT']

desired_titles = [
    "TLT_Policy Compliance Report_Windows 2019_Domain_Server_Monthly",
    "TLT_Policy Compliance Report_Windows 2019_Member Server_Monthly",
    "TLT_Policy Compliance Report_Windows 2016_Member Server_Monthly",
    "TLT_Policy Compliance Report_Windows 10_Monthly"
]

# These dictionaries will hold the latest report for each title
latest_2019_Domain = None
latest_2019_Member = None
latest_2016_Member = None
latest_10 = None

for report in reports:
    title = report.get('TITLE', 'Default Title')
    
    # Check for Windows 2019 Member Server Monthly report
    if title == "TLT_Policy Compliance Report_Windows 2019_Member Server_Monthly":
        if not latest_2019_Member or report['LAUNCH_DATETIME'] > latest_2019_Member['LAUNCH_DATETIME']:
            latest_2019_Member = report

    # Check for Windows 2019_Domain Server Monthly
    elif title == "TLT_Policy Compliance Report_Windows 2019_Domain_Server_Monthly":
        if not latest_2019_Domain or report['LAUNCH_DATETIME'] > latest_2019_Domain['LAUNCH_DATETIME']:
            latest_2019_Domain = report

    # Check for Windows 2016 Member Server Monthly report
    elif title == "TLT_Policy Compliance Report_Windows 2016_Member Server_Monthly":
        if not latest_2016_Member or report['LAUNCH_DATETIME'] > latest_2016_Member['LAUNCH_DATETIME']:
            latest_2016_Member = report
            
    # Check for Windows 10 Monthly report
    elif title == "TLT_Policy Compliance Report_Windows 10_Monthly":
        if not latest_10 or report['LAUNCH_DATETIME'] > latest_10['LAUNCH_DATETIME']:
            latest_10 = report

# Collect the latest reports in a list
latest_reports = [report for report in [latest_2019_Member, latest_2019_Domain , latest_2016_Member, latest_10] if report]

# print(json.dumps(latest_reports, indent=4))

# Extract the IDs and construct the download URLs
list_download_urls = [download_url + report['ID'] for report in latest_reports]

for report, window_report_url in zip(latest_reports, list_download_urls):
    response = requests.post(window_report_url, headers=headers)
    
    if response.status_code == 200:
        # Sanitize title for a filename: replace spaces with underscores, remove special characters
        title_sanitized = report['TITLE'].replace(' ', '_').replace('/', '_')
        file_name = f"{title_sanitized}_latest.csv"
        
        with open(file_name, 'wb') as file:
            file.write(response.content)
        print(f"Report saved to: {file_name}")
    else:
        print(f"Failed to download report from {window_report_url}. Status code: {response.status_code}")
