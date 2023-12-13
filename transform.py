import csv
from collections import defaultdict
from datetime import datetime

# Function for the first script
# def process_first_file(input_file_name):
#     input_file = input_file_name
#     temp_output = input_file.replace("latest", "temp")
#     final_output = input_file.replace("latest", "final")
    
#     # Process the file as in the first script
#     with open(input_file, "r") as csv_input:
#         csv_reader = csv.reader(csv_input)
#         rows_to_keep = []
#         next_row_to_keep = None
#         for row in csv_reader:
#             if "passed" in row or "Failures" in row:
#                 rows_to_keep.append(row)
#                 next_row_to_keep = next(csv_reader, None)
#                 if next_row_to_keep:
#                     rows_to_keep.append(next_row_to_keep)

#         remaining_rows = [row for row in csv_reader if row not in rows_to_keep]

#     with open(temp_output, "w", newline="") as csv_output:
#         csv_writer = csv.writer(csv_output)
#         csv_writer.writerows(remaining_rows)
#         csv_writer.writerows(rows_to_keep)

#     desired_headers = ['Technologies', 'Controls', 'Assets', 'Control Instances', 'Passed', 'Failures'] #, 'Passed', 'Failures'
#     desired_headers_displays = ['Technologies', 'Controls', 'Assets', 'Control Instances']

#     with open(temp_output, 'r', newline='') as temp_input, open(final_output, 'w', newline='') as output_file:
#         reader = csv.DictReader(temp_input)
#         writer = csv.DictWriter(output_file, fieldnames=desired_headers)
#         writer.writeheader()
#         for row in reader:
#             filtered_row = {header: row[header] for header in desired_headers}
#             writer.writerow(filtered_row)

# def process_first_file(input_file_name):
#     # ... (previous code remains unchanged) ...
#     input_file = input_file_name
#     temp_output = input_file.replace("latest", "temp")
#     final_output = input_file.replace("latest", "final")

#     # Process the file as in the first script
#     with open(input_file, "r") as csv_input:
#         csv_reader = csv.reader(csv_input)
#         rows_to_keep = []
#         next_row_to_keep = None
#         for row in csv_reader:
#             if "passed" in row or "Failures" in row:
#                 rows_to_keep.append(row)
#                 next_row_to_keep = next(csv_reader, None)
#                 if next_row_to_keep:
#                     rows_to_keep.append(next_row_to_keep)

#         remaining_rows = [row for row in csv_reader if row not in rows_to_keep]

#     with open(temp_output, "w", newline="") as csv_output:
#         csv_writer = csv.writer(csv_output)
#         csv_writer.writerows(remaining_rows)
#         csv_writer.writerows(rows_to_keep)

#     desired_headers = ['Technologies', 'Controls', 'Assets', 'Control Instances', 'Passed', 'Failures']
#     desired_headers_displays = ['Technologies', 'Controls', 'Assets', 'Control Instances']

#     # Modify values in 'Passed' and 'Failures' columns to extract only the numerical part before writing to the final output file
#     with open(temp_output, 'r', newline='') as temp_input, open(final_output, 'w', newline='') as output_file:
#         reader = csv.DictReader(temp_input)
#         writer = csv.DictWriter(output_file, fieldnames=desired_headers)
#         writer.writeheader()
#         for row in reader:
#             filtered_row = {header: row[header] if header not in ['Passed', 'Failures'] else row[header].replace('%', '').split('(')[0].strip() for header in desired_headers}
#             writer.writerow(filtered_row)

def process_first_file(input_file_name):
    input_file = input_file_name
    temp_output = input_file.replace("latest", "temp")
    final_output = input_file.replace("latest", "final")
    pass_fail_output = input_file.replace("latest", "a_passfail")

    # Process the file as in the first script
    with open(input_file, "r") as csv_input:
        csv_reader = csv.reader(csv_input)
        rows_to_keep = []
        next_row_to_keep = None
        for row in csv_reader:
            if "passed" in row or "Failures" in row:
                rows_to_keep.append(row)
                next_row_to_keep = next(csv_reader, None)
                if next_row_to_keep:
                    rows_to_keep.append(next_row_to_keep)

        remaining_rows = [row for row in csv_reader if row not in rows_to_keep]

    with open(temp_output, "w", newline="") as csv_output:
        csv_writer = csv.writer(csv_output)
        csv_writer.writerows(remaining_rows)
        csv_writer.writerows(rows_to_keep)

    desired_headers = ['Technologies', 'Controls', 'Assets', 'Control Instances', 'Passed', 'Failures']

    # Extract columns for the first CSV file
    with open(temp_output, 'r', newline='') as temp_input, open(final_output, 'w', newline='') as output_file:
        reader = csv.DictReader(temp_input)
        writer = csv.DictWriter(output_file, fieldnames=desired_headers[:4])  # First set of columns
        writer.writeheader()
        for row in reader:
            filtered_row = {header: row[header] for header in desired_headers[:4]}
            writer.writerow(filtered_row)

    # Extract columns for the second CSV file
    with open(temp_output, 'r', newline='') as temp_input, open(pass_fail_output, 'w', newline='') as output_file:
        reader = csv.DictReader(temp_input)
        writer = csv.DictWriter(output_file, fieldnames=desired_headers[4:])  # Second set of columns
        writer.writeheader()
        for row in reader:
            filtered_row = {header: row[header].replace('%', '').split('(')[0].strip() for header in desired_headers[4:]}
            writer.writerow(filtered_row)

def process_second_file(input_file_name):
    input_file = input_file_name
    temp_output = input_file.replace("latest", "temp2")
    final_output = input_file.replace("latest", "final2")
    
    with open(input_file, "r") as csv_input:
        csv_reader = csv.reader(csv_input)
        for row in csv_reader:
            if "Host IP" in row:
                host_ip_row = row
                break
        remaining_rows = [row for row in csv_reader]

    with open(temp_output, "w", newline="") as csv_output:
        csv_writer = csv.writer(csv_output)
        csv_writer.writerow(host_ip_row)
        csv_writer.writerows(remaining_rows)

    desired_headers = ["Host IP", "DNS Hostname","NetBIOS Hostname","Tracking Method","Operating System","OS CPE","NETWORK","Last Scan Date","Evaluation Date","Control ID","Technology","Control","Criticality Label","Criticality Value","Instance","Status","Remediation","Deprecated","Evidence","Cause of Failure","Qualys Host ID"]

    hostname_column = 'NetBIOS Hostname'
    if input_file_name == 'TLT_Policy_Compliance_Report_Windows_2019_Domain_Server_Monthly_latest.csv':
        hostname_column = 'DNS Hostname'

    counter = defaultdict(int)
    with open(temp_output, 'r', newline='') as temp_input, open(final_output, 'w', newline='') as output_file:
        reader = csv.DictReader(temp_input)
        for row in reader:
            counter[row[hostname_column]] += 1

        # Sort the counter by values (highest to lowest)
        sorted_counter = sorted(counter.items(), key=lambda x: x[1], reverse=True)
        
        writer = csv.DictWriter(output_file, fieldnames=['Hostname', 'Count of Control'])
        writer.writeheader()
        for hostname, count in sorted_counter:
            writer.writerow({'Hostname': hostname, 'Count of Control': count})

        # Add the Grand Total row
        grand_total = sum(counter.values())
        if grand_total != 0:
            writer.writerow({'Hostname': 'Grand Total', 'Count of Control': grand_total})



# List of input files
input_files = [
    "TLT_Policy_Compliance_Report_Windows_10_Monthly_latest.csv",
    "TLT_Policy_Compliance_Report_Windows_2016_Member_Server_Monthly_latest.csv",
    "TLT_Policy_Compliance_Report_Windows_2019_Domain_Server_Monthly_latest.csv",
    "TLT_Policy_Compliance_Report_Windows_2019_Member_Server_Monthly_latest.csv"
]

# Process each file with both functions
for input_file_name in input_files:
    process_first_file(input_file_name)
    process_second_file(input_file_name)
