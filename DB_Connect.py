import psycopg2
import csv

# Connect to the PostgreSQL database
conn = psycopg2.connect(database="powerbi", user="postgres", password="b45ebbed", host="127.0.0.1", port="5432")


# Create a cursor object
cur = conn.cursor()

# Open the CSV file and insert data into the table
with open('/Users/jiraphatromsaiyud/desktop/tlt_security/test.csv', 'r') as f:
    reader = csv.reader(f)
    header = next(reader)  # get the header row
    columns = ', '.join(f'"{col}"' for col in header)
    placeholders = ', '.join(['%s'] * len(header))  # create placeholders for the values
    for row in reader:
        values = [row[header.index(col)] for col in header]  # map the values to the correct columns
        query = f"INSERT INTO compliance ({columns}) VALUES ({placeholders})"
        cur.execute(query, values)


# Commit the changes and close the cursor and database connection
conn.commit()
cur.close()
conn.close()
