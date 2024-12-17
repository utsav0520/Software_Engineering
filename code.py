import pandas as pd
import random

# Load the Excel file
file_path = "DVA_Practical-1_uv.xlsx"
data = pd.read_excel(file_path)

# List of valid TOE Names
valid_toe_names = [
    "Airport System Planning",
    "Fundamentals of Disaster Management",
    "Introduction to Artificial Intelligence",
    "Energy Management System",
    "Modern Control System",
    "Introduction to Adaptive Signal Processing",
    "Wireless Mobile Communications",
    "Industrial Automation with PLC",
    "Basics of Building Automation",
    "Power Plant Engineering",
    "Heating, Ventilation and Airconditioning (HVAC)"
]

# Replace missing or invalid TOE Names for rows where SPI is between 6 and 10
data['TOE Name'] = data.apply(
    lambda row: random.choice(valid_toe_names) if pd.isna(row['TOE Name']) and 6 <= row['Current SPI'] <= 10 else row['TOE Name'],
    axis=1
)

# Save the updated file
updated_file_path = "updated_file.xlsx"
data.to_excel(updated_file_path, index=False)

print(f"File updated successfully and saved at {updated_file_path}")
