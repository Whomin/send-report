import smtplib
from quickchart import QuickChart
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os
import pandas as pd

def send_email_with_attachments(subject, body, to_email, cc_emails, file_paths, smtp_details):
    from_email = smtp_details['email']
    from_password = smtp_details['password']
    smtp_server = smtp_details['server']
    smtp_port = smtp_details['port']
    use_tls = smtp_details.get('use_tls', False)

    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['CC'] = ', '.join(cc_emails)  # Join multiple CC emails with a comma
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'html'))

    for file_path in file_paths:
        with open(file_path, 'rb') as attachment_file:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment_file.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f"attachment; filename= {file_path.split('/')[-1]}")
            msg.attach(part)

    recipients = to_email.split(",") + cc_emails  # Combine both TO and CC recipients

    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        if use_tls:
            server.starttls()
        server.login(from_email, from_password)
        server.sendmail(from_email, recipients, msg.as_string())
        server.quit()
        print("Email sent successfully!")
    except Exception as e:
        print(f"Failed to send email: {e}")

# File paths of the excel
file_paths = [
    "TLT_Policy_Compliance_Report_Windows_2019_Member_Server_Monthly_latest.xlsx",
    "TLT_Policy_Compliance_Report_Windows_2019_Member_Server_Monthly_ByHost_ByConTrol.xlsx",
    "TLT_Policy_Compliance_Report_Windows_2019_Domain_Server_Monthly_latest.xlsx",
    "TLT_Policy_Compliance_Report_Windows_2019_Domain_Server_Monthly_ByHost_ByConTrol.xlsx",
    "TLT_Policy_Compliance_Report_Windows_2016_Member_Server_Monthly_latest.xlsx",
    "TLT_Policy_Compliance_Report_Windows_2016_Member_Server_Monthly_ByHost_ByConTrol.xlsx",
    "TLT_Policy_Compliance_Report_Windows_10_Monthly_latest.xlsx",
    "TLT_Policy_Compliance_Report_Windows_10_Monthly_ByHost_ByConTrol.xlsx"
]

email_pass = os.environ.get('EMAIL_PASS', 'Default Value')
smtp_details = {
    'email': 'gotnotify@got.co.th',
    'password': 'fnGh%LS*F9F9QH',
    'server': 'smtp.office365.com',
    'port': 587,   # Often 587 for TLS or 465 for SSL
    'use_tls': True  # Set this to False if you don't want to use starttls
}

# smtp_details = {
#     'email': 'jitradbo@tlt.co.th',
#     'password': '0b9ifk[6Ppno*11',
#     'server': 'smtp.office365.com',
#     'port': 587,   # Often 587 for TLS or 465 for SSL
#     'use_tls': True  # Set this to False if you don't want to use starttls
# }

import csv
import datetime

what = ["TLT_Policy_Compliance_Report_Windows_2019_Member_Server_Monthly_ByHost_ByConTrol.xlsx","TLT_Policy_Compliance_Report_Windows_2019_Domain_Server_Monthly_ByHost_ByConTrol.xlsx","TLT_Policy_Compliance_Report_Windows_2016_Member_Server_Monthly_ByHost_ByConTrol.xlsx","TLT_Policy_Compliance_Report_Windows_10_Monthly_ByHost_ByConTrol.xlsx"]
passfail = ['TLT_Policy_Compliance_Report_Windows_10_Monthly_a_passfail.csv','TLT_Policy_Compliance_Report_Windows_2016_Member_Server_Monthly_a_passfail.csv','TLT_Policy_Compliance_Report_Windows_2019_Member_Server_Monthly_a_passfail.csv', 'TLT_Policy_Compliance_Report_Windows_2019_Domain_Server_Monthly_a_passfail.csv']

def is_number(n):
    try:
        float(n)  # For integer and float numbers
        return True
    except ValueError:
        return False

def csv_to_html_table(csv_filename):
    if csv_filename in passfail :
        with open(csv_filename, 'r') as file:
            reader = csv.DictReader(file)
            #headers = next(reader)
            passed_values = []
            failures_values = []
            
            for row in reader:
                passed_values.append(int(row['Passed']))
                failures_values.append(int(row['Failures']))
                
            total = sum(passed_values) + sum(failures_values)
            passed_percent = "{:.2f}".format((sum(passed_values) / total) * 100)
            failures_percent = "{:.2f}".format((sum(failures_values) / total) * 100)
            
            labels = ['Passed', 'Failures']
            sizes = [passed_percent, failures_percent]
            chart_data = ','.join(str(size) for size in sizes)
            
            QuickChartURL = f"https://quickchart.io/chart?c={{type:'doughnut',data:{{labels:{labels},datasets:[{{data:[{chart_data}],backgroundColor:['ForestGreen', 'IndianRed']}}]}},options:{{plugins:{{datalabels:{{color:'white'}}}},elements:{{arc:{{borderColor:'transparent',borderWidth: 0}}}}}}}}&width:200,height:200"
            
            # Create the HTML content with the embedded pie chart from QuickChart
            html_content = f'''
            <html>
                <body>
                    <img src="{QuickChartURL}" alt="Pie Chart" class="center" width="340" height="200">
                </body>
            </html>
            '''
            return html_content
        
    if csv_filename in what :
        
        data = pd.read_excel(csv_filename, sheet_name=[0, 1])
        
        tables = []  # List to store HTML tables for each sheet
        
        # Get the sheet names from the Excel file
        sheet_names = pd.ExcelFile(csv_filename).sheet_names
        
        for sheet_index in data:
            # Get the dataframe for the corresponding sheet index
            sheet_data = data[sheet_index]
            
            # Get the sheet name corresponding to the current index
            sheet_name = sheet_names[sheet_index]
            
            # Create a header with the sheet name
            tables.append(f"<h3>{sheet_name}</h3>")
            
            # Convert DataFrame to HTML table
            table = '<table border="1">'
            
            # Header row with the color
            header_color = '#9bddff'
            table += f'<tr style="background-color: {header_color};">'
            for header in sheet_data.columns:
                table += f'<th>{header}</th>'
            table += '</tr>'
            
            for idx, row in sheet_data.iterrows():
                # Alternating row colors
                if "final2" in csv_filename and idx == len(data) - 1:  # This is the last row and filename contains "final2"
                    row_color = header_color
                else:
                    row_color = '#e0ffff'  # Other rows color
                table += f'<tr style="background-color: {row_color};">'
                for cell in row:
                    cell = '' if pd.isnull(cell) else cell
                    # Determine text alignment
                    align = 'center' if str(cell).isdigit() else 'left'
                    if idx == len(sheet_data) - 1:  # This is the last row
                        table += f'<td style="text-align: {align};"><b>{cell}</b></td>'
                    else:
                        table += f'<td style="text-align: {align};">{cell}</td>'
                table += '</tr>'
            
            table += '</table>'
            tables.append(table)  # Store HTML table for each sheet

        # Combine HTML tables into a single string separated by "<br/>"
        combined_tables = "<br/>".join(tables)
        
        return combined_tables
        
    else :
        with open(csv_filename, 'r') as file:
            reader = csv.reader(file)
            headers = next(reader)

            # Check if there's data beyond the headers
            data_rows = list(reader)
            if not data_rows:
                return ''

            table = '<table border="1">'
            # Header row with the color #9bddff
            header_color = '#9bddff'
            table += f'<tr style="background-color: {header_color};">'
            for header in headers:
                table += f'<th>{header}</th>'
                #print(header)
            table += '</tr>'
            
            for idx, row in enumerate(data_rows):
                # if "ByHost_ByConTrol" in csv_filename and idx == len(data_rows) - 1:  # This is the last row and filename contains "final2"
                #     row_color = header_color
                # else:
                row_color = '#e0ffff'  # Other rows color
                table += f'<tr style="background-color: {row_color};">'
                for cell in row:
                    # Determine text alignment
                    align = 'center' if is_number(cell) else 'left'
                    if idx == len(data_rows) - 1:  # This is the last row
                        table += f'<td style="text-align: {align};"><b>{cell}</b></td>'
                    else:
                        table += f'<td style="text-align: {align};">{cell}</td>'
                table += '</tr>'
            table += '</table>'
            return table
    
            # table = '<table style="border-collapse: collapse; width: 100%;">' #'<table border="1">'
            # # Header row with the color
            # header_color = '#9bddff'
            # table += f'<tr style="background-color: {header_color};">'
            # for header in headers:
            #     table += f'<th style="padding: 8px; border: 1px solid #ddd;">{header}</th>'
            # table += '</tr>'
            # for idx, row in enumerate(data_rows):
            #     # Alternating row colors
            #     row_color = '#f9f9f9' if idx % 2 == 0 else '#ffffff'
            #     table += f'<tr style="background-color: {row_color};">'
            #     for cell in row:
            #         # Determine text alignment
            #         align = 'center' if cell.isdigit() else 'left'
            #         if idx == len(data_rows) - 1:  # This is the last row
            #             table += f'<td style="padding: 8px; border: 1px solid #ddd; text-align: {align};"><b>{cell}</b></td>'
            #         else:
            #             table += f'<td style="padding: 8px; border: 1px solid #ddd; text-align: {align};">{cell}</td>'
            #     table += '</tr>'
            # table += '</table>'
            # return table

        

today = datetime.datetime.today()
formatted_date = today.strftime("%d %B %Y")

sections = { 
    'Windows 10': ['TLT_Policy_Compliance_Report_Windows_10_Monthly_final.csv', 'TLT_Policy_Compliance_Report_Windows_10_Monthly_a_passfail.csv', 'TLT_Policy_Compliance_Report_Windows_10_Monthly_final2.csv',"TLT_Policy_Compliance_Report_Windows_10_Monthly_ByHost_ByConTrol.xlsx"],
    'Windows 2016': ['TLT_Policy_Compliance_Report_Windows_2016_Member_Server_Monthly_final.csv','TLT_Policy_Compliance_Report_Windows_2016_Member_Server_Monthly_a_passfail.csv' , 'TLT_Policy_Compliance_Report_Windows_2016_Member_Server_Monthly_final2.csv',"TLT_Policy_Compliance_Report_Windows_2016_Member_Server_Monthly_ByHost_ByConTrol.xlsx"],
    'Windows 2019': ['TLT_Policy_Compliance_Report_Windows_2019_Member_Server_Monthly_final.csv','TLT_Policy_Compliance_Report_Windows_2019_Member_Server_Monthly_a_passfail.csv' , 'TLT_Policy_Compliance_Report_Windows_2019_Member_Server_Monthly_final2.csv',"TLT_Policy_Compliance_Report_Windows_2019_Member_Server_Monthly_ByHost_ByConTrol.xlsx"],
    'Windows 2019 Domain': ['TLT_Policy_Compliance_Report_Windows_2019_Domain_Server_Monthly_final.csv', 'TLT_Policy_Compliance_Report_Windows_2019_Domain_Server_Monthly_a_passfail.csv', 'TLT_Policy_Compliance_Report_Windows_2019_Domain_Server_Monthly_final2.csv',"TLT_Policy_Compliance_Report_Windows_2019_Domain_Server_Monthly_ByHost_ByConTrol.xlsx"]
}

content = ''

for section, files in sections.items():
    tables = ""
    for file in files:
        table_content = csv_to_html_table(file)
        if table_content:
            tables += table_content + "<br/>"
    if tables:
        content += f"<h3>{section}</h3>{tables}"

html_content = f"""
<html>
    <body>
        <h2><u>Update Baseline Windows as of {formatted_date}</u></h2>
        {content}
    </body>
</html>
"""

# print(html_content)

send_email_with_attachments(
    "Baseline Windows", 
    html_content , 
    #"papon@got.co.th",
    #["Jiraphat.g@got.co.th"], 
    "mootum251@gmail.com",
    ["Jiraphat.g@got.co.th"],
    file_paths, 
    smtp_details
)

# send_email_with_attachments(
#     "Baseline Windows", 
#     html_content , 
#     "itinfra@tlt.co.th",
#     ["pradist_k@tlt.co.th", "kijnipat_s@tlt.co.th", "kalunyu_s@tlt.co.th", "phornmesa_k@tlt.co.th", "jitrada_b@tlt.co.th"], 
#     file_paths, 
#     smtp_details
# )
