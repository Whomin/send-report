import csv

# Define the input and output file names
input_file = "Compliance_Report_TLT_Policy_Compliance_Report_Windows_10_Monthly_tfs_st2_20230206.csv"
with open(input_file, "r") as csv_input:
    csv_reader = csv.reader(csv_input)
    # Set the name of the output file to the value of all_rows[0][1] followed by the string "test"
    # Read all the rows of the input file into a list
    all_rows = list(csv_reader)
    date_part, time_part = all_rows[0][1].split(" at ")
    output_file = f"{date_part.replace('/', '_')}Compliance_Report_Windows 10.csv"
# Open the input file and create a CSV reader object
with open(input_file, "r") as csv_input:
    csv_reader = csv.reader(csv_input)
    
    # Find the row that contains "Host IP" and save it
    for row in csv_reader:
        if "Host IP" in row:
            host_ip_row = row
            break
    
    # Create a list of the remaining rows
    remaining_rows = [row for row in csv_reader]
    
# Write the remaining rows to a new CSV file, with "Host IP" row at the top
with open(output_file, "w", newline="") as csv_output:
    csv_writer = csv.writer(csv_output)
    csv_writer.writerow(host_ip_row)
    csv_writer.writerows(remaining_rows)

# import csv

# # Define the input and output file names
# input_file = "Compliance_Report_TLT_Policy_Compliance_Report_Windows_10_Monthly_tfs_ks7_20230223.csv"
# output_file = "test02.csv"

# # Open the input file and create a CSV reader object
# with open(input_file, "r") as csv_input:
#     csv_reader = csv.reader(csv_input)
    
#     # Find the rows that contain "passed", "Failures", or "Error" and the row immediately following each of them
#     rows_to_keep = []
#     next_row_to_keep = None
#     for row in csv_reader:
#         if "passed" in row or "Failures" in row or "Error" in row:
#             rows_to_keep.append(row)
#             next_row_to_keep = next(csv_reader, None)
#             if next_row_to_keep:
#                 rows_to_keep.append(next_row_to_keep)
    
#     # Create a list of the remaining rows
#     remaining_rows = [row for row in csv_reader if row not in rows_to_keep]
    
# # Write the remaining rows to a new CSV file, followed by the rows to keep
# with open(output_file, "w", newline="") as csv_output:
#     csv_writer = csv.writer(csv_output)
#     csv_writer.writerows(remaining_rows)
#     csv_writer.writerows(rows_to_keep)
