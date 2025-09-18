import pandas as pd
from datetime import datetime

# Sample parks data in the exact format expected
parks_data = [
    {"name": "Anzac Park", "suburb": "Ipswich Central", "latitude": -27.614892, "longitude": 152.759412, "size_sqm": 15000, "difficulty": 3, "priority": 3, "last_mowed": "2024-01-15"},
    {"name": "Lions Park", "suburb": "West Ipswich", "latitude": -27.618234, "longitude": 152.752567, "size_sqm": 8500, "difficulty": 2, "priority": 2, "last_mowed": "2024-01-10"},
    {"name": "Heritage Park", "suburb": "Booval", "latitude": -27.608456, "longitude": 152.785234, "size_sqm": 45000, "difficulty": 4, "priority": 3, "last_mowed": "2023-12-28"},
    {"name": "Queens Park", "suburb": "East Ipswich", "latitude": -27.612034, "longitude": 152.765123, "size_sqm": 6500, "difficulty": 2, "priority": 3, "last_mowed": "2024-01-20"},
    {"name": "Rotary Park", "suburb": "North Ipswich", "latitude": -27.605123, "longitude": 152.758234, "size_sqm": 12000, "difficulty": 3, "priority": 2, "last_mowed": "2024-01-08"},
    {"name": "Pioneer Park", "suburb": "Bundamba", "latitude": -27.598456, "longitude": 152.772345, "size_sqm": 35000, "difficulty": 4, "priority": 3, "last_mowed": "2024-01-05"},
    {"name": "Memorial Park", "suburb": "Goodna", "latitude": -27.600123, "longitude": 152.745234, "size_sqm": 28000, "difficulty": 4, "priority": 4, "last_mowed": "2023-12-30"},
    {"name": "Central Park", "suburb": "Redbank", "latitude": -27.595234, "longitude": 152.875123, "size_sqm": 85000, "difficulty": 5, "priority": 3, "last_mowed": "2024-01-12"},
    {"name": "Victory Park", "suburb": "Raceview", "latitude": -27.630234, "longitude": 152.780123, "size_sqm": 9500, "difficulty": 3, "priority": 3, "last_mowed": "2024-01-18"},
    {"name": "Jacaranda Park", "suburb": "Yamanto", "latitude": -27.645123, "longitude": 152.790234, "size_sqm": 7200, "difficulty": 2, "priority": 3, "last_mowed": "2024-01-14"},
    {"name": "Riverside Park", "suburb": "Leichhardt", "latitude": -27.580234, "longitude": 152.730123, "size_sqm": 18000, "difficulty": 3, "priority": 2, "last_mowed": "2024-01-07"},
    {"name": "Discovery Park", "suburb": "One Mile", "latitude": -27.590123, "longitude": 152.740234, "size_sqm": 22000, "difficulty": 3, "priority": 4, "last_mowed": "2024-01-16"},
    {"name": "Eucalyptus Park", "suburb": "Silkstone", "latitude": -27.620234, "longitude": 152.730123, "size_sqm": 14500, "difficulty": 3, "priority": 3, "last_mowed": "2023-12-25"},
    {"name": "Federation Park", "suburb": "Brassall", "latitude": -27.610123, "longitude": 152.765234, "size_sqm": 11000, "difficulty": 3, "priority": 3, "last_mowed": "2024-01-11"},
    {"name": "Unity Park", "suburb": "Sadliers Crossing", "latitude": -27.585234, "longitude": 152.755123, "size_sqm": 32000, "difficulty": 4, "priority": 3, "last_mowed": "2024-01-03"},
    {"name": "Wattle Park", "suburb": "Ipswich Central", "latitude": -27.615123, "longitude": 152.759823, "size_sqm": 5500, "difficulty": 2, "priority": 2, "last_mowed": "2024-01-19"},
    {"name": "Botanical Park", "suburb": "West Ipswich", "latitude": -27.618723, "longitude": 152.752123, "size_sqm": 16500, "difficulty": 3, "priority": 4, "last_mowed": "2024-01-09"},
    {"name": "Adventure Park", "suburb": "Booval", "latitude": -27.608823, "longitude": 152.785567, "size_sqm": 4200, "difficulty": 1, "priority": 1, "last_mowed": "2024-01-22"},
    {"name": "Peace Park", "suburb": "East Ipswich", "latitude": -27.612456, "longitude": 152.765456, "size_sqm": 19000, "difficulty": 3, "priority": 3, "last_mowed": "2024-01-06"},
    {"name": "Community Park", "suburb": "North Ipswich", "latitude": -27.605456, "longitude": 152.758567, "size_sqm": 3800, "difficulty": 1, "priority": 3, "last_mowed": "2024-01-17"},
    {"name": "Environmental Park", "suburb": "Bundamba", "latitude": -27.598723, "longitude": 152.772678, "size_sqm": 25000, "difficulty": 4, "priority": 4, "last_mowed": "2024-01-04"},
    {"name": "Centenary Park", "suburb": "Goodna", "latitude": -27.600456, "longitude": 152.745567, "size_sqm": 13500, "difficulty": 3, "priority": 3, "last_mowed": "2024-01-13"},
    {"name": "Orion Park", "suburb": "Redbank", "latitude": -27.595567, "longitude": 152.875456, "size_sqm": 7800, "difficulty": 2, "priority": 2, "last_mowed": "2024-01-21"},
    {"name": "Family Park", "suburb": "Raceview", "latitude": -27.630567, "longitude": 152.780456, "size_sqm": 10500, "difficulty": 3, "priority": 3, "last_mowed": "2024-01-02"},
    {"name": "Banksia Park", "suburb": "Yamanto", "latitude": -27.645456, "longitude": 152.790567, "size_sqm": 6800, "difficulty": 2, "priority": 3, "last_mowed": "2024-01-23"}
]

# Create DataFrame
df = pd.DataFrame(parks_data)

# Ensure proper data types
df['latitude'] = df['latitude'].astype(float)
df['longitude'] = df['longitude'].astype(float)
df['size_sqm'] = df['size_sqm'].astype(int)
df['difficulty'] = df['difficulty'].astype(int)
df['priority'] = df['priority'].astype(int)

# Convert last_mowed to proper date format
df['last_mowed'] = pd.to_datetime(df['last_mowed'], errors='coerce')

print("Creating Excel file...")
print(f"DataFrame shape: {df.shape}")
print(f"Column names: {list(df.columns)}")
print(f"Data types:\n{df.dtypes}")

# Create Excel file with proper formatting
with pd.ExcelWriter('ipswich_parks.xlsx', engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='Parks', index=False)
    
    # Get workbook and worksheet for formatting
    workbook = writer.book
    worksheet = writer.sheets['Parks']
    
    # Format headers
    from openpyxl.styles import Font, PatternFill, Alignment
    
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="2D572C", end_color="2D572C", fill_type="solid")
    
    for cell in worksheet[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
    
    # Auto-adjust column widths
    for column in worksheet.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                cell_length = len(str(cell.value))
                if cell_length > max_length:
                    max_length = cell_length
            except:
                pass
        adjusted_width = min(max_length + 2, 25)
        worksheet.column_dimensions[column_letter].width = adjusted_width

print("✅ Excel file 'ipswich_parks.xlsx' created successfully!")
print("\nFirst few rows:")
print(df.head())
print(f"\nTotal parks: {len(df)}")
print(f"Suburbs: {df['suburb'].nunique()}")
print(f"Size range: {df['size_sqm'].min():,} - {df['size_sqm'].max():,} sqm")

# Also create a simple CSV version
df_csv = df.copy()
df_csv['last_mowed'] = df_csv['last_mowed'].dt.strftime('%Y-%m-%d')
df_csv.to_csv('ipswich_parks.csv', index=False)
print("✅ CSV file 'ipswich_parks.csv' also created!")

# Show what the upload system should see
print("\n" + "="*50)
print("DEBUG INFO FOR TROUBLESHOOTING:")
print("="*50)
print("Column names (case-sensitive):")
for i, col in enumerate(df.columns):
    print(f"  {i}: '{col}' (type: {df[col].dtype})")

print("\nSample data:")
for i, row in df.head(3).iterrows():
    print(f"Row {i}:")
    for col in df.columns:
        print(f"  {col}: {repr(row[col])}")
    print()