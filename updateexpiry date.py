import pandas as pd
from datetime import datetime

# Get today's date in YYYYMMDD format
today_date = datetime.now().strftime('%Y%m%d')

# Define file paths
nsb_file = f'NSBXTPLSH_{today_date}.csv'
xn_file = 'XN Available Stock Report_ADMIN_20250825034640810.xlsx'

# Read the two excel files
try:
    nsb_df = pd.read_csv(nsb_file)
    xn_df = pd.read_excel(xn_file)
except FileNotFoundError as e:
    print(f"Error: {e}. Please ensure the files are in the correct directory.")
    exit()



# Concatenate ITEMDESCRIPTION and COMBINEPACKING with " - " in both dataframes
nsb_df['combined'] = nsb_df['ITEMDESCRIPTION'].astype(str) + '- ' + nsb_df['COMBINEPACKING'].astype(str)
nsb_df['combined'] = nsb_df['combined'].astype(str).str.replace(' ', '')
xn_file['Description'] = xn_file['Description'].astype(str).str.replace(' ', '')



# Perform the XLOOKUP-like operation using pandas merge
# Merge the two dataframes on the 'combined' column to get the 'ExpiryDate'
#merged_df = pd.merge(nsb_df, xn_df[['combined', 'ExpiryDate']], on='combined', how='left')
merged_df = pd.merge(xn_df, nsb_df[['combined', 'ExpiryDate']], on='combined', how='left')

# The merged_df now contains the NSBXTPLSH data with an added 'ExpiryDate' column from the XN file
# You can now save this updated dataframe to a new CSV or Excel file
merged_df.to_csv(f'NSBXTPLSH_with_expiry_{today_date}.csv', index=False)

print(f"Successfully retrieved expiry dates and saved the updated file as 'NSBXTPLSH_with_expiry_{today_date}.csv'.")
print("\nFirst 5 rows of the updated dataframe:")
print(merged_df.head())