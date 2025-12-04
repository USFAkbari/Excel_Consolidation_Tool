import pandas as pd
from pathlib import Path

# Define paths
resource_folder = Path("resorce")
output_folder = Path("output")

# Create output folder if it doesn't exist
output_folder.mkdir(exist_ok=True)

# Get all Excel files from resource folder
excel_files = list(resource_folder.glob("*.xlsx")) + list(resource_folder.glob("*.xls"))

print(f"Found {len(excel_files)} Excel files to process")

# Read all Excel files and concatenate
all_dataframes = []
for file in excel_files:
    try:
        df = pd.read_excel(file)
        all_dataframes.append(df)
        print(f"Read {file.name}: {len(df)} rows")
    except Exception as e:
        print(f"Error reading {file.name}: {e}")

# Concatenate all dataframes
if all_dataframes:
    combined_df = pd.concat(all_dataframes, ignore_index=True)
    print(f"\nTotal rows after concatenation: {len(combined_df)}")
    
    # Group by name and phone, sum the count column
    # The Poyesh column will count how many times each name+phone combination appears
    grouped = combined_df.groupby(['name', 'phone']).agg({
        'count': 'sum',
    }).reset_index()
    
    # Count the number of duplicate rows (how many times each name+phone appears)
    poyesh_counts = combined_df.groupby(['name', 'phone']).size().reset_index(name='Poyesh')
    
    # Merge the Poyesh column with the grouped data
    result_df = grouped.merge(poyesh_counts, on=['name', 'phone'], how='left')
    
    # Sort by Poyesh descending (or by count descending - I'll sort by Poyesh as it represents duplicates)
    result_df = result_df.sort_values('Poyesh', ascending=False)
    
    # Reorder columns: name, phone, count, Poyesh
    result_df = result_df[['name', 'phone', 'count', 'Poyesh']]
    
    # Save to Excel
    output_file = output_folder / "All_poyesh.xlsx"
    result_df.to_excel(output_file, index=False)
    print(f"\nSaved All_poyesh.xlsx with {len(result_df)} rows")
    print(f"Output saved to: {output_file}")
else:
    print("No Excel files found or all files failed to read")

