import sqlite3
import pandas as pd

# Connect to the database
conn = sqlite3.connect('cement_delivery.db')

# Read the data into a pandas DataFrame
df = pd.read_sql_query("SELECT * FROM delivery_reports ORDER BY date", conn)

# Close the connection
conn.close()

# Export to CSV
df.to_csv('cement_delivery_data.csv', index=False)

print(f"Exported {len(df)} records to cement_delivery_data.csv")
print("\nColumn descriptions:")
print("- date: Report date")
print("- short: Short amount in KG")
print("- excess: Excess amount in KG")
print("- per_bag_short_excess: Per bag short/excess value")
print("- email_subject: Subject of the source email")
print("- email_received: When the email was received")