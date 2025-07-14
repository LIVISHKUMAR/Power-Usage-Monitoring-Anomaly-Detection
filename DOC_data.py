import pandas as pd
import matplotlib.pyplot as plt

# Load Excel file
file_path = 'power_data.xlsx'  # Replace with your file
df = pd.read_excel(file_path)

# Convert to datetime
df['Date Time'] = pd.to_datetime(df['Date Time'])
df = df.sort_values('Date Time')

# Extract Date
df['Date'] = df['Date Time'].dt.date

# Classify ON/OFF status
df['Status'] = df['Power'].apply(lambda x: 'ON' if x >= 21 else 'OFF')
df['Previous_Status'] = df['Status'].shift(1)
df['Transition'] = df['Previous_Status'] + ' → ' + df['Status']

# --- Opening Times: First OFF→ON transition per day
openings = df[df['Transition'] == 'OFF → ON'].groupby('Date').first().reset_index()
openings['Opening Time'] = openings['Date Time'].dt.strftime('%I:%M %p')  # AM/PM format

# --- Closing Times: Last ON→OFF transition per day
closings = df[df['Transition'] == 'ON → OFF'].groupby('Date').last().reset_index()
closings['Closing Time'] = closings['Date Time'].dt.strftime('%I:%M %p')  # AM/PM format

# --- Power Drops Below 21
df_drops = df[df['Power'] < 21][['Date Time', 'Power']].copy()
df_drops['Drop Time'] = df_drops['Date Time'].dt.strftime('%I:%M %p')  # AM/PM format
df_drops['Date'] = df_drops['Date Time'].dt.date

# Save to Excel
with pd.ExcelWriter('shop_power_summary.xlsx', engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='Full Data', index=False)
    openings[['Date', 'Opening Time']].to_excel(writer, sheet_name='Opening Times', index=False)
    closings[['Date', 'Closing Time']].to_excel(writer, sheet_name='Closing Times', index=False)
    df_drops[['Date', 'Drop Time', 'Power']].to_excel(writer, sheet_name='Power Drops < 21', index=False)

# Plotting
plt.figure(figsize=(14, 6))
plt.plot(df['Date Time'], df['Power'], label='Power Reading', color='blue')
plt.scatter(df_drops['Date Time'], df_drops['Power'], color='red', label='Power < 21', s=15)
plt.title('Power Meter Readings with Drops')
plt.xlabel('Date Time')
plt.ylabel('Power')
plt.legend()
plt.grid(True)
plt.tight_layout()
plt.savefig('power_readings_with_drops.png')
plt.show()
