import matplotlib.pyplot as plt
import pandas as pd
import csv

with open('TLT_Policy_Compliance_Report_Windows_10_Monthly_final.csv', 'r') as file:
    csv_reader = csv.DictReader(file)
    for row in csv_reader:
        passed_value = int(row['Passed'])
        failures_value = int(row['Failures'])

# Data for the pie chart
labels = ['Passed', 'Failures']
sizes = [passed_value, failures_value]
colors = ['#66b3ff', '#ff9999']  # Colors for Passed and Failures sections

# Plotting the pie chart
plt.figure(figsize=(8, 6))
plt.pie(sizes, labels=labels, colors=colors, autopct='%0.01f%%',startangle=140 )
plt.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
plt.title('Passed vs Failures Distribution')

# Show the pie chart
plt.show()