import os
import pandas as pd

#The aim of this code is to clean div_roa, dropping variables that are not used, transforming the date variable into year and quarter

# Get the desktop path
desktop_path = os.path.expanduser("~/Desktop/Capstone/Final Data folder")
file_path = os.path.join(desktop_path, "div_roa.csv")

# Load the DataFrame
df = pd.read_csv(file_path)

# Drop 'adate' and 'qdate' columns
df = df.drop(['adate', 'qdate'], axis=1)

# Convert 'public_date' to datetime
df['public_date'] = pd.to_datetime(df['public_date'])

# Create 'year' variable
df['year'] = df['public_date'].dt.year

# Create 'quarter' variable
df['quarter'] = df['public_date'].dt.quarter

# Drop 'public_date' column
df = df.drop('public_date', axis=1)

# Remove duplicate rows
df = df.drop_duplicates()

# Save the updated DataFrame
output_path = os.path.join(desktop_path, "cleaned_roa.csv")
df.to_csv(output_path, index=False)

# Print information about the changes
print(f"Original shape: {df.shape}")
print(f"Shape after processing: {df.shape}")
print(f"Processed dataset saved to: {output_path}")