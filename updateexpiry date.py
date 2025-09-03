import pandas as pd
from datetime import datetime
import numpy as np
from fuzzywuzzy import fuzz


# Get today's date in YYYYMMDD format
today_date = datetime.now().strftime('%Y%m%d')
threshold = 80


# Define file paths
nsb_file = f'NSBXTPLSH_20250825.csv'
xn_file = 'XN Available Stock Report_ADMIN_20250825034640810.xlsx'


def find_best_match(search_term, description_list, threshold=80):
    """Find best matching description above threshold"""
    best_match = None
    best_score = 0

    for desc in description_list:
        # Using token_set_ratio for better matching of product descriptions
        score = fuzz.token_set_ratio(str(search_term).lower(), str(desc).lower())

        if score >= threshold and score > best_score:
            best_score = score
            best_match = desc

    return best_match, best_score

# Read the two excel files
try:
    nsb_df = pd.read_csv(nsb_file)
    xn_df = pd.read_excel(xn_file)
except FileNotFoundError as e:
    print(f"Error: {e}. Please ensure the files are in the correct directory.")
    exit()

nsb_df['new_column'] = np.where(
    nsb_df['ITEMDESCRIPTION'] == nsb_df['COMBINEPACKING'],
    nsb_df['ITEMDESCRIPTION'],
    nsb_df['ITEMDESCRIPTION'] + '-' + nsb_df['COMBINEPACKING']
)

results = []

for idx, row in nsb_df.iterrows():
    search_term = row['new_column']

    # Find best match in xn_file descriptions
    best_match, score = find_best_match(search_term, xn_df['Description'].tolist(), threshold)

    if best_match:
        # Get expiry date for the matched description
        expiry_date = xn_df[xn_df['Description'] == best_match]['ExpiryDate'].iloc[0]
        results.append({
            'original_index': idx,
            'new_column': search_term,
            'matched_description': best_match,
            'match_score': score,
            'expiry_date': expiry_date
        })
    else:
        # No match found, use today's date
        results.append({
            'original_index': idx,
            'new_column': search_term,
            'matched_description': None,
            'match_score': 0,
            'expiry_date': today_date
        })

result_df = pd.DataFrame(results)
# Create copy of original dataframe and add new columns
nsb_file_copy = nsb_file.copy()
nsb_file_copy['expiry_date'] = result_df['expiry_date'].values
nsb_file_copy['matched_description'] = result_df['matched_description'].values
nsb_file_copy['match_score'] = result_df['match_score'].values

# Display results
print("Results with fuzzy matching:")
print(nsb_file_copy[['new_column', 'matched_description', 'match_score', 'expiry_date']])

# Filter to see only successful matches
matches = nsb_file_copy[nsb_file_copy['match_score'] >= 80]
print(f"\nSuccessful matches: {len(matches)}")

# Filter to see items that got today's date (no match)
no_matches = nsb_file_copy[nsb_file_copy['match_score'] < 80]
print(f"No matches (using today's date): {len(no_matches)}")


# Enhanced version using multiple similarity methods for better matching
def find_best_match_enhanced(search_term, description_list, threshold=80):
    best_match = None
    best_score = 0

    for desc in description_list:
        # Try multiple similarity methods
        scores = [
            fuzz.ratio(str(search_term).lower(), str(desc).lower()),
            fuzz.partial_ratio(str(search_term).lower(), str(desc).lower()),
            fuzz.token_sort_ratio(str(search_term).lower(), str(desc).lower()),
            fuzz.token_set_ratio(str(search_term).lower(), str(desc).lower())
        ]

        # Take the maximum score from all methods
        max_score = max(scores)

        if max_score >= threshold and max_score > best_score:
            best_score = max_score
            best_match = desc

    return best_match, best_score


# If you want to use the enhanced matching instead, uncomment this section:
"""
# Enhanced matching results
enhanced_results = []

for idx, row in nsb_file.iterrows():
    search_term = row['new_column']
    best_match, score = find_best_match_enhanced(search_term, xn_file['Description'].tolist(), threshold)

    if best_match:
        expiry_date = xn_file[xn_file['Description'] == best_match]['ExpiryDate'].iloc[0]
    else:
        expiry_date = today

    enhanced_results.append({
        'original_index': idx,
        'new_column': search_term,
        'matched_description': best_match,
        'match_score': score,
        'expiry_date': expiry_date
    })

enhanced_result_df = pd.DataFrame(enhanced_results)
nsb_file_enhanced = nsb_file.copy()
nsb_file_enhanced['expiry_date'] = enhanced_result_df['expiry_date'].values
nsb_file_enhanced['matched_description'] = enhanced_result_df['matched_description'].values
nsb_file_enhanced['match_score'] = enhanced_result_df['match_score'].values

print("Enhanced matching results:")
print(nsb_file_enhanced[['new_column', 'matched_description', 'match_score', 'expiry_date']])