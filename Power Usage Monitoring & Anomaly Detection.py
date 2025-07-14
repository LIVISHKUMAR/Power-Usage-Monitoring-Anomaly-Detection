import pandas as pd

# Load your data (adjust file path and format as needed)
df = pd.read_csv('power_meter_data.csv', parse_dates=['DateTime'])

# Sort by DateTime just in case
df = df.sort_values('DateTime')

# Extract date part for grouping
df['Date'] = df['DateTime'].dt.date

# Find minimum power level (assumed as baseline from chillers/freezers)
min_power = df['MeterValue'].min()

# Threshold above minimum to detect abnormal usage after closing (you can tune this)
threshold = min_power + 5  # example threshold, adjust based on your data

results = []

for date, group in df.groupby('Date'):
    # Find when power meter starts increasing (shop opening)
    diffs = group['MeterValue'].diff()
    
    # Opening time: first timestamp where diff > 0 (meter starts increasing)
    opening_idx = diffs[diffs > 0].index.min()
    if pd.isna(opening_idx):
        # No opening found for that day, skip
        continue
    opening_time = group.loc[opening_idx, 'DateTime']
    
    # Closing time: last timestamp before meter stops increasing (diff <= 0 after opening)
    after_opening = group.loc[opening_idx:]
    closing_idx = after_opening[after_opening['MeterValue'].diff() <= 0].index.min()
    
    if pd.isna(closing_idx):
        # If no closing time found, assume last record as closing
        closing_time = group.iloc[-1]['DateTime']
    else:
        closing_time = group.loc[closing_idx, 'DateTime']
    
    # Calculate opening hours and closing hours
    day_start = pd.Timestamp.combine(date, pd.Timestamp.min.time())
    day_end = pd.Timestamp.combine(date, pd.Timestamp.max.time())
    
    opening_hours = (closing_time - opening_time).seconds / 3600  # hours shop is open
    closing_hours = 24 - opening_hours
    
    # Check power values after closing time for abnormal usage
    after_close = group[group['DateTime'] > closing_time]
    abnormal = after_close['MeterValue'].max() > threshold
    
    abnormal_status = 'power level is normal' if not abnormal else 'equipment running after closing'
    
    results.append({
        'Date': date,
        'Opening Time': opening_time.time(),
        'Closing Time': closing_time.time(),
        'Opening Hours': f"{int(opening_hours)} hrs",
        'Closing Hours': f"{int(closing_hours)} hrs",
        'Abnormal Usage After Closing': abnormal_status
    })

# Convert results to DataFrame and export to Excel
output_df = pd.DataFrame(results)
output_df.to_excel('shop_power_usage_report.xlsx', index=False)

print("Report generated: shop_power_usage_report.xlsx")
